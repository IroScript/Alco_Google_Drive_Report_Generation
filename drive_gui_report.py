import datetime
import os
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Service Account and API logic
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
        _, ext = os.path.splitext(filename)
        if ext:
            return ext[1:].upper()
        return 'Unknown'

def generate_report(start_date_str, end_date_str, status_label, root):
    try:
        # Validate dates
        start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d")
        if start_date > end_date:
            messagebox.showerror("Error", "Start Date cannot be after End Date.")
            return
    except ValueError:
        messagebox.showerror("Error", "Please use YYYY-MM-DD format for dates.")
        return

    status_label.config(text="Querying Drive...", foreground="blue")
    root.update()

    service = get_drive_service()
    
    # Format RFC3339 times
    start_rfc = f"{start_date_str}T00:00:00Z"
    end_rfc = f"{end_date_str}T23:59:59Z"
    
    query = f"createdTime >= '{start_rfc}' and createdTime <= '{end_rfc}' and mimeType != 'application/vnd.google-apps.folder' and trashed = false"
    
    files = []
    page_token = None
    while True:
        try:
            results = service.files().list(
                q=query,
                pageSize=1000,
                pageToken=page_token,
                fields="nextPageToken, files(id, name, mimeType, createdTime, parents)",
            ).execute()
            files.extend(results.get('files', []))
            page_token = results.get('nextPageToken')
            if not page_token:
                break
        except Exception as e:
            status_label.config(text="Error occurred", foreground="red")
            messagebox.showerror("API Error", f"Failed to fetch files: {e}")
            return

    if not files:
        status_label.config(text="No files found", foreground="orange")
        messagebox.showinfo("No Files", "No files were found in the selected date range.")
        return

    status_label.config(text="Processing folders...", foreground="blue")
    root.update()

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

    folder_data = {}  # Key: folder_name, Value: {'file_types': set, 'qty': int}

    for file in files:
        parents = file.get('parents', [])
        parent_names = []
        for p in parents:
            p_name = get_folder_name(p)
            parent_names.append(p_name)
        
        parent_display = ", ".join(parent_names) if parent_names else "Root / Shared with me"
        
        # Extract clean extension/type in lowercase, e.g. 'pdf', 'png', 'jpeg'
        _, ext = os.path.splitext(file['name'])
        file_type = ext[1:].lower() if ext else ''
        if not file_type:
            # Fallback based on mimeType
            if 'pdf' in file['mimeType']:
                file_type = 'pdf'
            elif 'image' in file['mimeType']:
                # generic fallback
                if 'jpeg' in file['mimeType'] or 'jpg' in file['mimeType']:
                    file_type = 'jpeg'
                elif 'png' in file['mimeType']:
                    file_type = 'png'
                else:
                    file_type = 'image'
            elif 'spreadsheet' in file['mimeType'] or 'excel' in file['mimeType']:
                file_type = 'xlsx'
            else:
                file_type = 'unknown'
        
        # Normalize jpg -> jpeg if desired, or keep as is
        if file_type == 'jpg':
            file_type = 'jpeg'

        if parent_display not in folder_data:
            folder_data[parent_display] = {'file_types': set(), 'qty': 0}
            
        folder_data[parent_display]['file_types'].add(file_type)
        folder_data[parent_display]['qty'] += 1

    summary_list = []
    for folder, data in folder_data.items():
        sorted_types = sorted(list(data['file_types']))
        file_types_str = ", ".join(sorted_types)
        
        summary_list.append({
            'Folder Name': folder,
            'File Types': file_types_str,
            'Qty': data['qty']
        })
        
    df_summary = pd.DataFrame(summary_list)
    if not df_summary.empty:
        df_summary = df_summary.sort_values(by='Folder Name')
    else:
        df_summary = pd.DataFrame(columns=['Folder Name', 'File Types', 'Qty'])

    # File output path
    output_filename = f"drive_report_{start_date_str}_to_{end_date_str}.xlsx"
    output_path = os.path.abspath(os.path.join(os.path.dirname(__file__), output_filename))
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # Auto-adjust column widths
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for col in worksheet.columns:
                    max_len = max(len(str(val.value or '')) for val in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 15)
        
        status_label.config(text="Report generated!", foreground="green")
        messagebox.showinfo("Success", f"Report generated successfully:\n{output_filename}")
    except Exception as e:
        status_label.config(text="Save failed", foreground="red")
        messagebox.showerror("File Error", f"Failed to save Excel file: {e}")

# Build GUI
def build_gui():
    root = tk.Tk()
    root.title("Drive Reporter")
    root.geometry("380x280")
    root.resizable(False, False)
    
    # Modern styling
    style = ttk.Style(root)
    style.theme_use('clam')
    
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Title
    title_label = ttk.Label(main_frame, text="Google Drive Report Generator", font=("Arial", 12, "bold"))
    title_label.pack(pady=(0, 15))
    
    # Date Range Input
    dates_frame = ttk.LabelFrame(main_frame, text=" Date Range (YYYY-MM-DD) ", padding="10")
    dates_frame.pack(fill=tk.X, pady=(0, 15))
    
    # Start Date
    start_frame = ttk.Frame(dates_frame)
    start_frame.pack(fill=tk.X, pady=5)
    ttk.Label(start_frame, text="Start Date:", width=10).pack(side=tk.LEFT)
    start_entry = ttk.Entry(start_frame)
    start_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    # Default start date (7 days ago)
    seven_days_ago = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%Y-%m-%d")
    start_entry.insert(0, seven_days_ago)
    
    # End Date
    end_frame = ttk.Frame(dates_frame)
    end_frame.pack(fill=tk.X, pady=5)
    ttk.Label(end_frame, text="End Date:", width=10).pack(side=tk.LEFT)
    end_entry = ttk.Entry(end_frame)
    end_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    # Default end date (today)
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    end_entry.insert(0, today_str)
    
    # Status Label
    status_label = ttk.Label(main_frame, text="Ready", font=("Arial", 9, "italic"))
    status_label.pack(pady=5)
    
    # Button
    def on_generate():
        generate_report(start_entry.get().strip(), end_entry.get().strip(), status_label, root)
        
    btn_generate = ttk.Button(main_frame, text="Generate Report", command=on_generate)
    btn_generate.pack(pady=(5, 0), fill=tk.X)
    
    root.mainloop()

if __name__ == '__main__':
    build_gui()
