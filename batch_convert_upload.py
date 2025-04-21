import os
import shutil
import subprocess
import fitz
from datetime import datetime
import webbrowser

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ['https://www.googleapis.com/auth/drive.file']


def authenticate_drive():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)


def extract_folder_and_title_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = doc[0].get_text()
    doc.close()
    folder_name = "UnknownDate"
    title = "Unnamed"
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "Event Worksheet" in line and i + 1 < len(lines):
            date_line = lines[i + 1].strip()
            folder_name = date_line.replace("/", "-").replace(" ", "_")
        if "Event Title:" in line and i + 1 < len(lines):
            title = lines[i + 1].strip()
    return folder_name, title


def read_drive_folder_id(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Folder ID file not found: {file_path}")
    with open(file_path, 'r') as f:
        return f.read().strip()


def set_file_permission(service, file_id):
    permission = {
        'type': 'anyone',
        'role': 'reader'
    }
    service.permissions().create(fileId=file_id, body=permission).execute()


def upload_to_drive(service, file_path, parent_folder_id, subfolder_name):
    results = service.files().list(q=f"'{parent_folder_id}' in parents and name='{subfolder_name}' and mimeType='application/vnd.google-apps.folder'",
                                   spaces='drive', fields='files(id, name)').execute()
    folder = results.get('files')
    if folder:
        folder_id = folder[0]['id']
    else:
        file_metadata = {
            'name': subfolder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }
        folder = service.files().create(body=file_metadata, fields='id').execute()
        folder_id = folder.get('id')
        set_file_permission(service, folder_id)

    # Check for existing file and delete if found
    existing_files = service.files().list(q=f"'{folder_id}' in parents and name='{os.path.basename(file_path)}'",
                                          spaces='drive', fields='files(id)').execute().get('files', [])
    for file in existing_files:
        service.files().delete(fileId=file['id']).execute()

    file_metadata = {'name': os.path.basename(file_path), 'parents': [folder_id]}
    media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    uploaded = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    set_file_permission(service, uploaded.get('id'))
    print(f"Uploaded {file_path} to Google Drive folder '{subfolder_name}'.")
    webbrowser.open(uploaded.get('webViewLink'))


def main():
    input_dir = "./input_files"
    output_dir = "./converted_docs"
    processed_dir = "./processed_files"
    folder_id_file = "drive_folder_id.txt"

    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(processed_dir, exist_ok=True)

    drive_service = authenticate_drive()
    parent_drive_folder_id = read_drive_folder_id(folder_id_file)

    for filename in os.listdir(input_dir):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_dir, filename)
            folder_name, title = extract_folder_and_title_from_pdf(pdf_path)
            safe_title = title.replace("/", "_")
            docx_name = f"{safe_title}.docx"
            docx_path = os.path.join(output_dir, docx_name)

            print(f"Converting {filename} to {docx_name}...")
            subprocess.run(["python", "convert.py", "--pdf", pdf_path, "--out", docx_path], check=True)

            upload_to_drive(drive_service, docx_path, parent_drive_folder_id, folder_name)
            shutil.move(pdf_path, os.path.join(processed_dir, filename))
            print(f"Moved {filename} to processed folder.")


if __name__ == "__main__":
    main()
