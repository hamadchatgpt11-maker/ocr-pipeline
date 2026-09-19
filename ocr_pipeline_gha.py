import os
import re
import io
import requests
import fitz
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload, MediaIoBaseUpload

GOOGLE_SCRIPT_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbz4BT3ZC2Pnv2IMLLnZ4n4fY-9yVy4zSNhxTE8UsRx6Qc75lwDgXfIH7dEke27xAn94Jw/exec"
SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'service_account.json'

creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive = build('drive', 'v3', credentials=creds)


def find_folder(name, parent_id=None):
    q = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    res = drive.files().list(q=q, fields="files(id, name)").execute()
    files = res.get('files', [])
    return files[0]['id'] if files else None


def create_folder(name, parent_id=None):
    metadata = {'name': name, 'mimeType': 'application/vnd.google-apps.folder'}
    if parent_id:
        metadata['parents'] = [parent_id]
    return drive.files().create(body=metadata, fields='id').execute()['id']


def get_or_create_folder(name, parent_id=None):
    fid = find_folder(name, parent_id)
    return fid if fid else create_folder(name, parent_id)


def list_files(folder_id):
    q = f"'{folder_id}' in parents and trashed=false"
    res = drive.files().list(q=q, fields="files(id, name)", pageSize=1000).execute()
    return res.get('files', [])


def download_text(file_id):
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, drive.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return fh.getvalue().decode('utf-8')


def upload_text(folder_id, name, content, existing_file_id=None):
    media = MediaIoBaseUpload(io.BytesIO(content.encode('utf-8')), mimetype='text/plain')
    if existing_file_id:
        drive.files().update(fileId=existing_file_id, media_body=media).execute()
    else:
        drive.files().create(body={'name': name, 'parents': [folder_id]}, media_body=media, fields='id').execute()


def upload_image(folder_id, local_path, name):
    media = MediaFileUpload(local_path, mimetype='image/jpeg')
    drive.files().create(body={'name': name, 'parents': [folder_id]}, media_body=media, fields='id').execute()


def run_once():
    main_folder_id = get_or_create_folder("PDFTOOCR")
    files = list_files(main_folder_id)

    links_file = next((f for f in files if f['name'] == 'links.txt'), None)
    if not links_file:
        upload_text(main_folder_id, "links.txt", "# یہاں ہر لائن میں ایک PDF کا آن لائن لنک (URL) پیسٹ کریں\n")
        print("📁 links.txt نہیں ملی، بنا دی گئی۔")
        return

    content = download_text(links_file['id'])
    found_urls = re.findall(r'https?://[^\s]+', content)
    pdf_urls = [u for u in found_urls if not u.startswith('#')]

    if not pdf_urls:
        print("⏳ links.txt فی الحال خالی ہے۔")
        return

    pdf_url = pdf_urls[0]
    raw_name = pdf_url.split('/')[-1].split('?')[0]
    book_name = Path(raw_name).stem or "book_item_auto"
    book_folder_name = f"OCR_Book_{book_name}"

    print(f"📖 کتاب کا لنک: {pdf_url}")

    # ✅ Step 1: check kya OCR pehle hi complete ho chuka hai (docx PDFTOOCR mein maujood hai)
    docx_done = any(f['name'].endswith('.docx') and book_name in f['name'] for f in list_files(main_folder_id))
    if docx_done:
        print(f"🎉 '{book_name}' مکمل ہو چکی ہے — لنک ہٹایا جا رہا ہے۔")
        upload_text(main_folder_id, "links.txt", content.replace(pdf_url, ""), existing_file_id=links_file['id'])
        return

    # ✅ Step 2: book folder + images
    book_folder_id = find_folder(book_folder_name, main_folder_id)
    if not book_folder_id:
        book_folder_id = create_folder(book_folder_name, main_folder_id)

    book_files = list_files(book_folder_id)
    existing_images = [f for f in book_files if f['name'].lower().endswith(('.jpg', '.jpeg', '.png'))]

    if not existing_images:
        print(f"📥 کتاب ڈاؤن لوڈ ہو رہی ہے: {pdf_url}")
        temp_pdf_path = f"/tmp/{book_name}.pdf"
        r = requests.get(pdf_url, stream=True)
        r.raise_for_status()
        with open(temp_pdf_path, "wb") as f:
            for chunk in r.iter_content(8192):
                f.write(chunk)

        doc = fitz.open(temp_pdf_path)
        total_pages = len(doc)
        print(f"📄 کل صفحات: {total_pages}")

        for i in range(total_pages):
            pix = doc[i].get_pixmap(dpi=200)
            local_img = f"/tmp/page_{i + 1}.jpg"
            pix.save(local_img, jpg_quality=85)
            upload_image(book_folder_id, local_img, f"page_{i + 1}.jpg")
            os.remove(local_img)
            if (i + 1) % 50 == 0:
                print(f"   ...{i + 1}/{total_pages} اپلوڈ ہو چکیں")

        os.remove(temp_pdf_path)
        print("✅ تمام تصاویر اپلوڈ ہو گئیں۔")
        book_files = list_files(book_folder_id)

    # ✅ Step 3: agar abhi tak trigger nahi hua to trigger karo
    already_triggered = any(f['name'] == 'triggered.flag' for f in book_files)

    if not already_triggered:
        print("🔄 گوگل اسکرپٹ کو سگنل بھیجا جا رہا ہے...")
        trigger_url = f"{GOOGLE_SCRIPT_WEB_APP_URL}?action=start&book={book_name}"
        res = requests.get(trigger_url, timeout=60)
        print("📡 جواب:", res.text)
        upload_text(book_folder_id, "triggered.flag", "started")
    else:
        print(f"⏳ '{book_name}' کا OCR پہلے سے چل رہا ہے — اگلی run میں دوبارہ چیک ہوگا۔")


if __name__ == "__main__":
    run_once()
