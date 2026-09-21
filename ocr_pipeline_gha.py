"""
OCR pipeline for GitHub Actions (OAuth2 refresh-token auth, Google Drive API v3).

Flow (one book per run):
  1. Read links.txt from the PDFTOOCR folder in Google Drive
  2. Pick the first book that is not finished
  3. Download the PDF, render pages to JPEG (200 DPI, quality 85)
  4. Upload the images to OCR_Book_<name> in PDFTOOCR
  5. Signal Apps Script (?action=start&book=<name>)
  6. Write triggered.flag into the book folder
  7. Once .docx output is found -> remove that link from links.txt

Required env vars (GitHub secrets):
  GDRIVE_CLIENT_ID, GDRIVE_CLIENT_SECRET, GDRIVE_REFRESH_TOKEN
"""
import io
import os
import re
import sys
import tempfile
from urllib.parse import unquote, urlparse, quote

import requests

try:
    import pymupdf as fitz
except ImportError:  # older PyMuPDF
    import fitz

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ----------------------------- CONFIG ---------------------------------
APPS_SCRIPT_URL = os.environ.get(
    "APPS_SCRIPT_URL",
    "https://script.google.com/macros/s/"
    "AKfycbz4BT3ZC2Pnv2IMLLnZ4n4fY-9yVy4zSNhxTE8UsRx6Qc75lwDgXfIH7dEke27xAn94Jw/exec",
)
ROOT_FOLDER_NAME = "PDFTOOCR"
LINKS_FILE_NAME = "links.txt"
BOOK_FOLDER_PREFIX = "OCR_Book_"
FLAG_NAME = "triggered.flag"
OUTPUT_FOLDER_NAME = "Output Files"   # finished .docx files go here
DPI = 200
JPEG_QUALITY = 85
FOLDER_MIME = "application/vnd.google-apps.folder"
# ----------------------------------------------------------------------


def get_drive():
    creds = Credentials(
        token=None,
        refresh_token=os.environ["GDRIVE_REFRESH_TOKEN"].strip(),
        client_id=os.environ["GDRIVE_CLIENT_ID"].strip(),
        client_secret=os.environ["GDRIVE_CLIENT_SECRET"].strip(),
        token_uri="https://oauth2.googleapis.com/token",
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    try:
        creds.refresh(Request())
    except RefreshError as e:
        print("ERROR: could not refresh Google token:", e)
        print("If this says invalid_grant, the refresh token expired (Testing mode = 7 days).")
        print("Get a new refresh token from OAuth Playground and update GDRIVE_REFRESH_TOKEN.")
        sys.exit(1)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


# --------------------------- Drive helpers -----------------------------
def q_escape(s):
    return s.replace("\\", "\\\\").replace("'", "\\'")


def find_one(drive, name, parent_id=None, folder=False):
    q = f"name='{q_escape(name)}' and trashed=false"
    if folder:
        q += f" and mimeType='{FOLDER_MIME}'"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    res = drive.files().list(q=q, fields="files(id,name)", pageSize=10).execute(num_retries=5)
    files = res.get("files", [])
    return files[0]["id"] if files else None


def create_folder(drive, name, parent_id):
    body = {"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]}
    return drive.files().create(body=body, fields="id").execute(num_retries=5)["id"]


def list_children(drive, parent_id):
    items, token = [], None
    while True:
        res = drive.files().list(
            q=f"'{parent_id}' in parents and trashed=false",
            fields="nextPageToken, files(id,name,mimeType)",
            pageSize=1000,
            pageToken=token,
        ).execute(num_retries=5)
        items.extend(res.get("files", []))
        token = res.get("nextPageToken")
        if not token:
            return items


def upload_bytes(drive, folder_id, name, data, mime):
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=False)
    body = {"name": name, "parents": [folder_id]}
    return drive.files().create(body=body, media_body=media, fields="id").execute(num_retries=5)


def upload_text(drive, folder_id, name, text):
    return upload_bytes(drive, folder_id, name, text.encode("utf-8"), "text/plain")


def update_text(drive, file_id, text):
    media = MediaIoBaseUpload(io.BytesIO(text.encode("utf-8")), mimetype="text/plain", resumable=False)
    return drive.files().update(fileId=file_id, media_body=media).execute(num_retries=5)


def read_text(drive, file_id):
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk(num_retries=5)
    return buf.getvalue().decode("utf-8-sig", errors="replace")


# --------------------------- Book helpers ------------------------------
def parse_links(text):
    return [ln.strip() for ln in text.splitlines() if ln.strip().lower().startswith("http")]


def remove_link(drive, file_id, raw_text, url_to_remove):
    lines = raw_text.splitlines()
    new_lines = [ln for ln in lines if ln.strip() != url_to_remove.strip()]
    new_text = "\n".join(new_lines) + ("\n" if new_lines else "")
    update_text(drive, file_id, new_text)


def book_name_from_url(url):
    stem = os.path.splitext(unquote(os.path.basename(urlparse(url).path)))[0]
    return re.sub(r"[^A-Za-z0-9_\-]+", "_", stem).strip("_") or "book"


def book_state(drive, root_names, name):
    """Returns (state, folder_id): new | partial | running | done
    ✅ FIX: pehle .docx (finished) check hota hai, folder existence check ke pehle —
    warna book-folder delete hone ke baad ye ghalti se 'new' samajh leta tha."""
    finished = any(
        n.lower().endswith(".docx") and name.lower() in n.lower() for n in root_names
    )
    if finished:
        return "done", None

    folder_id = find_one(drive, BOOK_FOLDER_PREFIX + name, parent_id=None, folder=True)
    if not folder_id:
        return "new", None

    names = [f["name"] for f in list_children(drive, folder_id)]
    if any("completed" in n.lower() for n in names):
        return "done", folder_id

    if FLAG_NAME in names:
        return "running", folder_id

    return "partial", folder_id


def download_pdf(url, path):
    print("Downloading PDF:", url)
    with requests.get(url, stream=True, timeout=120) as r:
        r.raise_for_status()
        with open(path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                f.write(chunk)
    print("PDF size: %.1f MB" % (os.path.getsize(path) / 1e6))


def process_book(drive, root_id, url, name, folder_id):
    if not folder_id:
        folder_id = create_folder(drive, BOOK_FOLDER_PREFIX + name, root_id)
    existing = {f["name"] for f in list_children(drive, folder_id)}

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = os.path.join(tmp, "book.pdf")
        download_pdf(url, pdf_path)
        doc = fitz.open(pdf_path)
        total = len(doc)
        print("Pages:", total)
        zoom = DPI / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        for i in range(total):
            img_name = "page_%04d.jpg" % (i + 1)
            if img_name in existing:
                continue
            pix = doc[i].get_pixmap(matrix=matrix, colorspace=fitz.csRGB)
            data = pix.tobytes("jpeg", jpg_quality=JPEG_QUALITY)
            upload_bytes(drive, folder_id, img_name, data, "image/jpeg")
            if (i + 1) % 10 == 0 or i + 1 == total:
                print("Uploaded %d/%d" % (i + 1, total))
        doc.close()

    signal_url = "%s?action=start&book=%s" % (APPS_SCRIPT_URL, quote(name))
    print("Signalling Apps Script...")
    resp = requests.get(signal_url, timeout=180)
    print("Apps Script reply:", resp.status_code, resp.text[:300])
    resp.raise_for_status()

    upload_text(drive, folder_id, FLAG_NAME, "started")
    print("Done for this run:", name)


def run_once():
    drive = get_drive()

    root_id = find_one(drive, ROOT_FOLDER_NAME, folder=True)
    if not root_id:
        print("ERROR: folder '%s' not found in Drive" % ROOT_FOLDER_NAME)
        sys.exit(1)

    links_id = find_one(drive, LINKS_FILE_NAME, parent_id=root_id)
    if not links_id:
        print("ERROR: %s not found inside %s" % (LINKS_FILE_NAME, ROOT_FOLDER_NAME))
        sys.exit(1)

    raw_text = read_text(drive, links_id)
    links = parse_links(raw_text)
    print("Links found:", len(links))

    root_names = [f["name"] for f in list_children(drive, root_id)]
    out_id = find_one(drive, OUTPUT_FOLDER_NAME, parent_id=root_id, folder=True)
    if out_id:
        root_names += [f["name"] for f in list_children(drive, out_id)]

    for url in links:
        name = book_name_from_url(url)
        state, folder_id = book_state(drive, root_names, name)
        print("Book %s -> %s" % (name, state))

        if state == "done":
            print("Book already finished — links.txt se link hataya ja raha hai")
            remove_link(drive, links_id, raw_text, url)
            continue

        if state == "running":
            print("Previous book still being OCR'd, waiting for next run.")
            return

        process_book(drive, root_id, url, name, folder_id)
        return

    print("Nothing to do: all books are finished.")


if __name__ == "__main__":
    run_once()
