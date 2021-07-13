from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaIoBaseDownload
import argparse
import io
from pathlib import Path

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
]

# The ID of a sample document.
DOCUMENT_ID = "1cnThtFVGjls0Te0P91mQXp0KJBXeOJFikkmQScY9vpk"

template_keyword = "_________CompanyName________"

DOWNLOAD_PATH = "D:\Downloads"

def main():
    # get credentials from oauth or from cached
    creds = get_credentials()

    # create service library objects
    doc_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    # replace template with company
    replace_in_doc(doc_service, template_keyword, args.comp)

    # download files
    download_doc(drive_service)

    # revert back to template
    replace_in_doc(doc_service, args.comp, template_keyword)

def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def replace_in_doc(doc_service, target, replaceText):
    revision_id = get_revision_id(doc_service)
    update = (
        doc_service.documents()
        .batchUpdate(
            documentId=DOCUMENT_ID,
            body={
                "requests": [
                    {
                        "replaceAllText": {
                            "containsText": {
                                "matchCase": False,
                                "text": target,
                            },
                            "replaceText": replaceText,
                        }
                    }
                ],
                "writeControl": {"requiredRevisionId": revision_id},
            },
        )
        .execute()
    )
    print(update.get("replies"))


def get_revision_id(doc_service):
    # Retrieve the documents contents from the Docs service.
    document = doc_service.documents().get(documentId=DOCUMENT_ID).execute()
    revision_id = document.get("revisionId")
    return revision_id


def download_doc(drive_service):
    request = drive_service.files().export_media(
        fileId=DOCUMENT_ID, mimeType="application/pdf"
    )
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    file_path = os.path.join(DOWNLOAD_PATH, args.comp + ".pdf")
    print(file_path)
    with open(file_path, "wb+") as f:
        f.write(fh.getbuffer())


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("comp", type=str, help="company to replace")
    args = parser.parse_args()
    print(template_keyword, args.comp)
    main()
