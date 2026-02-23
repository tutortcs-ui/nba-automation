# =============================================================================
# drive_saver.py
# PURPOSE: Upload generated documents to a specific Google Drive folder
#          using a service account (no user login required).
#
# HOW IT WORKS:
#   - A service account is a special Google account that acts on your behalf
#   - You share your "NBA Documents" Drive folder with the service account email
#   - The service account can then upload files to that folder silently
#   - No browser login popup, no OAuth flow — fully automatic
#
# CREDENTIALS:
#   - Stored in Streamlit secrets (never in code or committed to GitHub)
#   - Accessed via st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]
#   - If secrets are missing, all functions return gracefully — app still works
#
# USAGE:
#   from core.drive_saver import save_to_drive, is_drive_configured
#
#   if is_drive_configured():
#       url = save_to_drive(file_bytes, "Analysis_PEC-702A.docx", folder_id)
# =============================================================================

import json
import io


def is_drive_configured() -> bool:
    """
    Check if Google Drive secrets are available in Streamlit.
    Returns True only if both the service account JSON and folder ID are set.
    This is called before any Drive operation so the app degrades gracefully
    when Drive is not configured.
    """
    try:
        import streamlit as st
        return (
            "GOOGLE_SERVICE_ACCOUNT_JSON" in st.secrets
            and "GDRIVE_FOLDER_ID" in st.secrets
        )
    except Exception:
        return False


def _get_drive_service():
    """
    Build and return an authenticated Google Drive API service object.
    Uses the service account credentials stored in Streamlit secrets.

    Returns None if credentials are missing or invalid.
    """
    try:
        import streamlit as st
        from google.oauth2 import service_account
        from googleapiclient.discovery import build

        # Load service account JSON from Streamlit secrets
        # st.secrets stores it as a string — parse it back to dict
        sa_info = st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]

        # Handle both string and dict formats
        if isinstance(sa_info, str):
            sa_dict = json.loads(sa_info)
        else:
            # Streamlit may auto-parse TOML — convert AttrDict to plain dict
            sa_dict = dict(sa_info)

        credentials = service_account.Credentials.from_service_account_info(
            sa_dict,
            scopes=["https://www.googleapis.com/auth/drive.file"],
            # drive.file scope: can only access files created by this app
            # More restrictive than drive scope — safer
        )

        service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        return service

    except Exception as e:
        print(f"  Drive: could not build service — {e}")
        return None


def save_to_drive(
    file_bytes: bytes,
    filename: str,
    folder_id: str = None,
) -> str | None:
    """
    Upload a file (as bytes) to a Google Drive folder.

    Args:
        file_bytes : raw bytes of the .docx file
        filename   : name to give the file in Drive e.g. "Analysis_PEC-702A.docx"
        folder_id  : Google Drive folder ID. If None, reads from st.secrets.
                     Get folder ID from the folder URL:
                     drive.google.com/drive/folders/THIS_PART_IS_THE_ID

    Returns:
        str  : shareable view URL of the uploaded file
        None : if upload failed (app continues without Drive save)
    """
    try:
        from googleapiclient.http import MediaIoBaseUpload
        import streamlit as st

        service = _get_drive_service()
        if service is None:
            return None

        # Use folder_id from argument, or fall back to Streamlit secret
        if folder_id is None:
            folder_id = st.secrets.get("GDRIVE_FOLDER_ID", None)
        if folder_id is None:
            print("  Drive: no folder ID configured")
            return None

        # File metadata — name and which folder it goes into
        file_metadata = {
            "name": filename,
            "parents": [folder_id],
        }

        # Wrap bytes in a stream for the API
        media = MediaIoBaseUpload(
            io.BytesIO(file_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            resumable=False,
        )

        # Upload the file
        uploaded = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink",   # only fetch what we need
        ).execute()

        file_id  = uploaded.get("id")
        view_url = uploaded.get("webViewLink", f"https://drive.google.com/file/d/{file_id}/view")

        print(f"  ✓ Saved to Drive: {filename} → {view_url}")
        return view_url

    except Exception as e:
        print(f"  Drive upload failed for {filename}: {e}")
        return None


def save_both_to_drive(
    analysis_bytes: bytes,
    atr_bytes: bytes,
    course_code: str,
) -> dict:
    """
    Upload both Analysis and ATR documents to Drive.
    Returns a dict with 'analysis_url' and 'atr_url' (either may be None on failure).

    This is the main function called from app.py.
    """
    try:
        import streamlit as st
        folder_id = st.secrets.get("GDRIVE_FOLDER_ID", None)
    except Exception:
        folder_id = None

    analysis_url = save_to_drive(
        analysis_bytes,
        f"Analysis_{course_code}.docx",
        folder_id,
    )
    atr_url = save_to_drive(
        atr_bytes,
        f"ATR_{course_code}.docx",
        folder_id,
    )

    return {
        "analysis_url": analysis_url,
        "atr_url":      atr_url,
    }
