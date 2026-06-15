import datetime
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build

def get_drive_service():
    # Look for credentials file in parent directory or current directory
    creds_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'alco-pharma-cf4b49e394bb.json'))
    if not os.path.exists(creds_path):
        creds_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'alco-pharma-cf4b49e394bb.json'))
    
    SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
    creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def main():
    service = get_drive_service()
    
    # Get start of today in UTC
    now = datetime.datetime.now(datetime.UTC)
    today_midnight = datetime.datetime(now.year, now.month, now.day, tzinfo=datetime.UTC)
    rfc3339_time = today_midnight.isoformat().replace('+00:00', 'Z')
    
    print(f"Querying files created or modified since: {rfc3339_time} (UTC)\n")
    
    # Query for files uploaded/modified today
    query = f"(createdTime >= '{rfc3339_time}' or modifiedTime >= '{rfc3339_time}') and trashed = false"
    try:
        results = service.files().list(
            q=query,
            pageSize=50,
            fields="nextPageToken, files(id, name, mimeType, createdTime, modifiedTime, parents)",
            orderBy="modifiedTime desc"
        ).execute()
        files = results.get('files', [])
    except Exception as e:
        print(f"Error querying files: {e}")
        files = []
    
    # If no files found for today, let's fetch the 15 most recently modified files as a sample
    if not files:
        print("No files created or modified today. Fetching the 15 most recent files/folders instead as a sample:\n")
        try:
            results = service.files().list(
                q="trashed = false",
                pageSize=15,
                fields="nextPageToken, files(id, name, mimeType, createdTime, modifiedTime, parents)",
                orderBy="modifiedTime desc"
            ).execute()
            files = results.get('files', [])
        except Exception as e:
            print(f"Error fetching recent files: {e}")
            return
        
    if not files:
        print("No files found in Google Drive.")
        return

    # Helper function to get parent name
    folder_cache = {}
    def get_folder_name(folder_id):
        if folder_id in folder_cache:
            return folder_cache[folder_id]
        try:
            folder_info = service.files().get(fileId=folder_id, fields="name").execute()
            name = folder_info.get('name')
            folder_cache[folder_id] = name
            return name
        except Exception:
            return f"Unknown ({folder_id})"

    print(f"{'Name':<40} | {'Type':<8} | {'Created Time':<20} | {'Modified Time':<20} | {'Parent Folders'}")
    print("-" * 125)
    for file in files:
        parents = file.get('parents', [])
        parent_names = ", ".join([get_folder_name(p) for p in parents]) if parents else "Root / Shared with me"
        
        is_folder = "Folder" if file['mimeType'] == 'application/vnd.google-apps.folder' else "File"
        created_time = file.get('createdTime', 'N/A')[:19].replace('T', ' ') if 'createdTime' in file else 'N/A'
        modified_time = file.get('modifiedTime', 'N/A')[:19].replace('T', ' ') if 'modifiedTime' in file else 'N/A'
            
        print(f"{file['name'][:40]:<40} | {is_folder:<8} | {created_time:<20} | {modified_time:<20} | {parent_names}")

if __name__ == '__main__':
    main()
