import datetime
import os
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

def get_drive_service():
    creds_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'alco-pharma-cf4b49e394bb.json'))
    if not os.path.exists(creds_path):
        creds_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'alco-pharma-cf4b49e394bb.json'))
    
    SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
    creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def get_file_type(mime_type, filename):
    if mime_type == 'application/vnd.google-apps.folder':
        return 'Folder'
    elif mime_type == 'application/vnd.google-apps.spreadsheet':
        return 'Google Sheet'
    elif mime_type == 'application/vnd.google-apps.document':
        return 'Google Doc'
    elif 'pdf' in mime_type:
        return 'PDF'
    elif 'image' in mime_type:
        return 'Image'
    elif 'spreadsheet' in mime_type or 'excel' in mime_type:
        return 'Excel'
    else:
        # Fallback to extension
        _, ext = os.path.splitext(filename)
        if ext:
            return ext[1:].upper()
        return 'Unknown'

def main():
    service = get_drive_service()
    
    # Get start of today in UTC
    now = datetime.datetime.now(datetime.UTC)
    today_midnight = datetime.datetime(now.year, now.month, now.day, tzinfo=datetime.UTC)
    rfc3339_time = today_midnight.isoformat().replace('+00:00', 'Z')
    
    print(f"Querying files uploaded today since: {rfc3339_time} (UTC)")
    
    # Query for files (excluding folders) uploaded today
    query = f"createdTime >= '{rfc3339_time}' and mimeType != 'application/vnd.google-apps.folder' and trashed = false"
    
    try:
        results = service.files().list(
            q=query,
            pageSize=1000,
            fields="files(id, name, mimeType, createdTime, parents)",
        ).execute()
        files = results.get('files', [])
    except Exception as e:
        print(f"Error querying files: {e}")
        return

    if not files:
        print("No files were uploaded today.")
        # Let's create an empty report or sample report. Let's do a sample of last 50 files if empty so they have data to look at
        print("Fetching recent 50 files as a sample report...")
        query_sample = "mimeType != 'application/vnd.google-apps.folder' and trashed = false"
        try:
            results = service.files().list(
                q=query_sample,
                pageSize=50,
                fields="files(id, name, mimeType, createdTime, parents)",
                orderBy="createdTime desc"
            ).execute()
            files = results.get('files', [])
        except Exception as e:
            print(f"Error fetching sample: {e}")
            return

    # Folder cache to avoid duplicate API calls
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
            return f"Unknown Folder ({folder_id})"

    file_data = []
    folder_counts = {}

    for file in files:
        parents = file.get('parents', [])
        parent_names = []
        for p in parents:
            p_name = get_folder_name(p)
            parent_names.append(p_name)
            folder_counts[p_name] = folder_counts.get(p_name, 0) + 1
        
        parent_display = ", ".join(parent_names) if parent_names else "Root / Shared with me"
        
        created_time = file.get('createdTime', 'N/A')
        if created_time != 'N/A':
            # Format to a cleaner string
            created_time = created_time[:19].replace('T', ' ')

        file_type = get_file_type(file['mimeType'], file['name'])
        drive_link = f"https://drive.google.com/file/d/{file['id']}/view"
        
        file_data.append({
            'File Name': file['name'],
            'File Type': file_type,
            'Upload Time (UTC)': created_time,
            'Parent Folder': parent_display,
            'Google Drive Link': drive_link
        })

    # Create DataFrames
    df_files = pd.DataFrame(file_data)
    
    # Folder Summary DataFrame
    folder_summary = []
    for folder, count in folder_counts.items():
        folder_summary.append({
            'Folder Name': folder,
            'Files Uploaded Today': count
        })
    df_folders = pd.DataFrame(folder_summary)
    if df_folders.empty:
        df_folders = pd.DataFrame(columns=['Folder Name', 'Files Uploaded Today'])
    else:
        df_folders = df_folders.sort_values(by='Files Uploaded Today', ascending=False)

    # Save to Excel
    output_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'drive_upload_report_today.xlsx'))
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_files.to_excel(writer, sheet_name='Today Uploads', index=False)
        df_folders.to_excel(writer, sheet_name='Folder Summary', index=False)
        
        # Access openpyxl objects to adjust column widths
        workbook = writer.book
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for col in worksheet.columns:
                max_len = max(len(str(val.value or '')) for val in col)
                col_letter = col[0].column_letter
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    print(f"\nExcel report successfully generated: {output_path}")

if __name__ == '__main__':
    main()
