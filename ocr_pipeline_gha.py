"""
OCR pipeline for GitHub Actions (OAuth2 refresh-token auth, Google Drive API v3).

Fully autonomous flow (one book per run, safe to run on a schedule forever):
  1. Read links.txt from the PDFTOOCR folder in Google Drive
  2. Pick the first book that is not finished and not already in progress
  3. Download the PDF, render pages to JPEG (200 DPI, quality 85)
  4. Upload the images to OCR_Book_<name> in PDFTOOCR
  5. Mark the book "in progress" with a ROOT-level marker file
     (this marker does NOT live inside OCR_Book_<name>, so it survives
     even if the Apps Script deletes that folder once OCR finishes)
  6. Signal Apps Script (?action=start&book=<name>)
  7. On a later run, once the finished .docx is found in Drive ->
     remove the in-progress marker AND remove that link from links.txt

Why the marker lives at the root:
  The Apps Script deletes OCR_Book_<name> as soon as OCR completes.
  If "is this book still running?" were judged by that folder's
  existence, a run that happens to execute in the gap between
  "folder deleted" and "docx saved" would wrongly conclude the book
  was never started, and would re-download and re-OCR it from
  scratch. The root-level marker is untouched by that cleanup, so
  the book is correctly treated as "still running" until its .docx
  actually shows up.

Required env vars (GitHub secrets):
  GDRIVE_CLIENT_ID, GDRIVE_CLIENT_SECRET, GDRIVE_REFRESH_TOKEN
"""
import io
import os
import re
import sys
import tempfile
from datetime import datetime, timedelta, timezone
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
OUTPUT_FOLDER_NAME = "Output Files"        # finished .docx files may live here
INPROGRESS_PREFIX = "INPROGRESS_"          # root-level marker, e.g. INPROGRESS_MAQAM_E_MAHFOOZ.flag
INPROGRESS_SUFFIX = ".flag"
STALE_AFTER_HOURS = 6                      # if a marker is older than this with no .docx yet, retry
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
            fields="nextPageToken, files(id,name,mimeType,modifiedTime)",
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


def trash_file(drive, file_id):
    drive.files().update(fileId=file_id, body={"trashed": True}).execute(num_retries=5)


# --------------------------- Book helpers ------------------------------
def parse_links(text):
    return [ln.strip() for ln in text.splitlines() if ln.strip().lower().startswith("http")]


def remove_link(text, url_to_remove):
    lines = text.splitlines()
    new_lines = [ln for ln in lines if ln.strip() != url_to_remove.strip()]
    return "\n".join(new_lines) + ("\n" if new_lines else "")


def book_name_from_url(url):
    stem = os.path.splitext(unquote(os.path.basename(urlparse(url).path)))[0]
    return re.sub(r"[^A-Za-z0-9_\-]+", "_", stem).strip("_") or "book"


def norm(s):
    """Loose match key: ignore case and any non-alphanumeric characters,
    so 'OCR - OCR_Book_MAQAM_E_MAHFOOZ.docx' matches 'MAQAM_E_MAHFOOZ'."""
    return re.sub(r"[^a-z0-9]+", "", s.lower())


def parse_drive_time(ts):
    return datetime.fromisoformat(ts.replace("Z", "+00:00"))


def book_state(name, root_files):
    """
    Decide what to do with a book using ONLY root-level Drive listing
    (root_files: list of {name, mimeType, modifiedTime} in PDFTOOCR,
    already merged with Output Files if that folder exists).

    Returns one of: "done", "running", "stale", "new"
    - done:    a finished .docx exists -> remove link, clean up marker
    - running: an INPROGRESS marker exists and is fresh -> skip this run
    - stale:   an INPROGRESS marker exists but is too old -> retry
    - new:     nothing exists yet -> start processing
    """
    nname = norm(name)

    finished = any(
        f["name"].lower().endswith(".docx") and nname in norm(f["name"])
        for f in root_files
    )
    if finished:
        return "done"

    marker = next(
        (f for f in root_files if f["name"] == INPROGRESS_PREFIX + name + INPROGRESS_SUFFIX),
        None,
    )

    if marker:
        age = datetime.now(timezone.utc) - parse_drive_time(marker["modifiedTime"])
        if age > timedelta(hours=STALE_AFTER_HOURS):
            return "stale"
        return "running"

    return "new"


def download_pdf(url, path):
    print("Downloading PDF:", url)
    with requests.get(url, stream=True, timeout=120) as r:
        r.raise_for_status()
        with open(path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                f.write(chunk)
    print("PDF size: %.1f MB" % (os.path.getsize(path) / 1e6))


def process_book(drive, root_id, url, name):
    """Download + upload images, write the root in-progress marker, signal Apps Script."""
    folder_id = find_one(drive, BOOK_FOLDER_PREFIX + name, parent_id=None, folder=True)
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

    # Write/refresh the ROOT-level in-progress marker BEFORE signalling Apps Script,
    # so even if this run is interrupted right after signalling, the next run sees
    # "running" instead of starting over.
    marker_name = INPROGRESS_PREFIX + name + INPROGRESS_SUFFIX
    existing_marker_id = find_one(drive, marker_name, parent_id=root_id)
    if existing_marker_id:
        update_text(drive, existing_marker_id, "started")
    else:
        upload_text(drive, root_id, marker_name, "started")

    signal_url = "%s?action=start&book=%s" % (APPS_SCRIPT_URL, quote(name))
    print("Signalling Apps Script...")
    resp = requests.get(signal_url, timeout=180)
    print("Apps Script reply:", resp.status_code, resp.text[:300])
    resp.raise_for_status()

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

    root_files = list_children(drive, root_id)
    out_id = find_one(drive, OUTPUT_FOLDER_NAME, parent_id=root_id, folder=True)
    if out_id:
        root_files += list_children(drive, out_id)

    current_text = raw_text
    changed = False

    for url in links:
        name = book_name_from_url(url)
        state = book_state(name, root_files)
        print("Book %s -> %s" % (name, state))

        if state == "done":
            print("Book already finished — links.txt se link hataya ja raha hai, marker saaf kiya ja raha hai")
            marker_id = find_one(drive, INPROGRESS_PREFIX + name + INPROGRESS_SUFFIX, parent_id=root_id)
            if marker_id:
                trash_file(drive, marker_id)
            current_text = remove_link(current_text, url)
            changed = True
            continue

        if state == "running":
            print("Book abhi OCR ho rahi hai (in-progress marker fresh hai) — agle run ka intezar")
            if changed:
                update_text(drive, links_id, current_text)
            return

        # state is "new" or "stale" -> (re)start processing this book, then stop for this run
        if state == "stale":
            print("In-progress marker %s+ ghante purana tha, dobara try kiya ja raha hai" % STALE_AFTER_HOURS)

        if changed:
            update_text(drive, links_id, current_text)
        process_book(drive, root_id, url, name)
        return

    if changed:
        update_text(drive, links_id, current_text)

    print("Nothing to do: all books are finished.")


if __name__ == "__main__":
    run_once()
