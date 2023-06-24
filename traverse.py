from googleapiclient.discovery import build
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'googleKey.json'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)

def traverse_directory(root_dir_id):
    """Traverse the given Google Drive directory and return a list of all subdirectories."""
    folders = []
    page_token = None
    while True:
        response = drive_service.files().list(q=f"'{root_dir_id}' in parents and mimeType='application/vnd.google-apps.folder'",
                                              spaces='drive',
                                              fields='nextPageToken, files(id, name)',
                                              pageToken=page_token).execute()
        for file in response.get('files', []):
            # Append a tuple of (id, name) instead of just the name
            folders.append((file.get('id'), file.get('name')))
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break
    return folders
