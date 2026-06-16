"""Upload Forex_Insights.xlsx to a Google Drive folder after each run.

Authentication uses a Google Service Account JSON key stored as the
GOOGLE_SERVICE_ACCOUNT_JSON GitHub secret (the full JSON content as a string).

Requires:
    pip install google-api-python-client google-auth

Set these in GitHub repo secrets:
    GOOGLE_SERVICE_ACCOUNT_JSON  — full JSON content of the service account key
    GOOGLE_DRIVE_FOLDER_ID       — ID of the Drive folder to upload into
                                   (get it from the folder URL: .../folders/<ID>)
"""

import json
import os

EXCEL_FILE = "Forex_Insights.xlsx"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SCOPES = ["https://www.googleapis.com/auth/drive.file"]


def _build_drive_service():
    """Build and return an authenticated Google Drive service client."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
    if not sa_json:
        raise EnvironmentError("GOOGLE_SERVICE_ACCOUNT_JSON secret is not set")

    sa_info = json.loads(sa_json)
    credentials = service_account.Credentials.from_service_account_info(
        sa_info, scopes=SCOPES
    )
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def _find_existing_file(service, folder_id: str) -> str | None:
    """Return the Drive file ID if Forex_Insights.xlsx already exists in folder."""
    query = (
        f"name='{EXCEL_FILE}' and "
        f"'{folder_id}' in parents and "
        f"trashed=false"
    )
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    return files[0]["id"] if files else None


def upload_to_drive() -> None:
    """Upload (or update) Forex_Insights.xlsx in the configured Google Drive folder."""
    folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "")
    if not folder_id:
        print("  SKIP Drive upload: GOOGLE_DRIVE_FOLDER_ID not set")
        return

    if not os.path.exists(EXCEL_FILE):
        print(f"  SKIP Drive upload: {EXCEL_FILE} not found")
        return

    try:
        from googleapiclient.http import MediaFileUpload

        service = _build_drive_service()
        media = MediaFileUpload(EXCEL_FILE, mimetype=MIME_XLSX, resumable=False)
        existing_id = _find_existing_file(service, folder_id)

        if existing_id:
            # Update existing file in-place (keeps sharing settings and links stable)
            service.files().update(
                fileId=existing_id,
                media_body=media,
                supportsAllDrives=True,
            ).execute()
            print(f"  OK Drive: updated existing file (id={existing_id})")
        else:
            # Create new file in the specified folder
            metadata = {"name": EXCEL_FILE, "parents": [folder_id]}
            uploaded = (
                service.files()
                .create(body=metadata, media_body=media, fields="id", supportsAllDrives=True)
                .execute()
            )
            print(f"  OK Drive: uploaded new file (id={uploaded.get('id')})")

    except Exception as exc:
        print(f"  WARN Drive upload failed: {exc}")