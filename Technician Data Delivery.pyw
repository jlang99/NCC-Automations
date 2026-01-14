import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys, os
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import requests

# Google API Imports
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
import io

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import EMAILS, CREDS, TRACKER_MAPPING, CB_ISSUES_SHEET, get_google_credentials


def send_email(recipient_email, site_name, attachment_paths):
    """Sends an email with the requested data as PDF attachments."""
    sender_email = EMAILS['NCC Desk']
    smtp_password = CREDS['shiftsumEmail']
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    date_str = datetime.now().strftime("%Y-%m-%d")

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ', '.join(recipient_email) if isinstance(recipient_email, list) else recipient_email
    msg['Subject'] = f"Requested Data {date_str} for {site_name}"

    # HTML body
    body = f"""
    <html>
        <body>
            <p>Hello,</p>
            <p>Attached is the requested data for {site_name} for {date_str}.</p>
            <p>This is an accurate example of the final output, sent to technicians scheduled for repairs/troubleshooting.<br>
            All feed back is welcomed!</p>

            <p>Thank you,<br>NCC Automation</p>
        </body>
    </html>
    """
    msg.attach(MIMEText(body, 'html'))

    # Attach the files
    for attachment_path in attachment_paths:
        try:
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(attachment_path)}")
            msg.attach(part)
        except FileNotFoundError:
            messagebox.showerror("Error", f"Attachment file not found at {attachment_path}")
            return

    # Send the email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, smtp_password)
            server.send_message(msg)
        messagebox.showinfo("Success", f"Email sent successfully to {recipient_email}!")
    except Exception as e:
        messagebox.showerror("Email Error", f"Failed to send email: {e}")

def download_and_send():
    """Handles the main logic of downloading the sheet and sending the email."""
    technician = tech_var.get()
    site = site_var.get()
    get_tracker_map = tracker_map_var.get()
    get_cb_list = cb_underperformance_var.get()

    if not technician or not site:
        messagebox.showwarning("Selection Missing", "Please select both a technician and a site.")
        return

    if not get_tracker_map and not get_cb_list:
        messagebox.showwarning("Selection Missing", "Please select at least one report to send.")
        return

    recipient_email = EMAILS.get(technician)
    if not recipient_email:
        messagebox.showerror("Error", f"Could not find email for {technician}.")
        return

    creds = get_google_credentials()
    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    attachment_paths = []
    date_str = datetime.now().strftime("%Y-%m-%d")
    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

    # 1. Handle Tracker Map
    if get_tracker_map:
        tracker_spreadsheet_id = TRACKER_MAPPING.get(site, {}).get('sheet_id')

        if not tracker_spreadsheet_id:
            messagebox.showerror("Error", f"Could not find Tracker Map spreadsheet ID for {site}.")
        else:
            file_path = os.path.join(download_dir, f"{site}_tracker_map_{date_str}.pdf")
            try:
                request = drive_service.files().export_media(fileId=tracker_spreadsheet_id,
                                                             mimeType='application/pdf')
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    print(f"Download Tracker Map for {site}: {int(status.progress() * 100)}%.")

                with open(file_path, 'wb') as f:
                    f.write(fh.getvalue())
                attachment_paths.append(file_path)

            except HttpError as error:
                messagebox.showerror("API Error", f"An error occurred downloading the Tracker Map: {error}")

    # 2. Handle CB Underperformance List
    if get_cb_list:
        cb_spreadsheet_id = CB_ISSUES_SHEET
        file_path = os.path.join(download_dir, f"{site}_CB_Underperformance_{date_str}.pdf")

        try:
            # Find the gid of the sheet matching the site name
            sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=cb_spreadsheet_id).execute()
            gid = None
            for s in sheet_metadata.get('sheets', []):
                if s.get('properties', {}).get('title') == site:
                    gid = s.get('properties', {}).get('sheetId')
                    break
            
            if gid is not None:
                # Use requests to download the specific sheet as a PDF
                headers = {'Authorization': 'Bearer ' + creds.token}
                url = f'https://docs.google.com/spreadsheets/d/{cb_spreadsheet_id}/export?format=pdf&gid={gid}'
                
                print(f"Downloading CB Underperformance PDF for {site}...")
                res = requests.get(url, headers=headers)
                res.raise_for_status()
                
                with open(file_path, 'wb') as f:
                    f.write(res.content)
                attachment_paths.append(file_path)
                print(f"Successfully downloaded CB Underperformance PDF for {site}.")

            else:
                messagebox.showerror("Error", f"Sheet '{site}' not found in the CB Underperformance workbook.")
        except HttpError as error:
            messagebox.showerror("API Error", f"An error occurred with Google Sheets API: {error}")
        except requests.exceptions.RequestException as error:
            messagebox.showerror("Download Error", f"An error occurred downloading the CB Underperformance PDF: {error}")

    # 3. Send the email with all collected attachments
    if attachment_paths:
        send_email(recipient_email, site, attachment_paths)
    else:
        messagebox.showinfo("No Files", "No files were downloaded to be sent. Check for previous errors.")

# --- GUI Setup ---
root = tk.Tk()
root.title("Technician Data Delivery")

main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Technician Dropdown
ttk.Label(main_frame, text="Select Technician:").grid(column=0, row=0, sticky=tk.W, pady=5)
tech_var = tk.StringVar()
tech_dropdown = ttk.Combobox(main_frame, textvariable=tech_var, state="readonly")
tech_dropdown['values'] = sorted([
    "Administrators + NCC","Isaac Million", "Jon Wieber", "Joseph Lang", "Parker Wilson", 
    "Thomas Rhodes", "Thorne Locklear", "Zach Duggan"
])
tech_dropdown.grid(column=1, row=0, sticky=(tk.W, tk.E), pady=5, padx=5)

# Site Dropdown
ttk.Label(main_frame, text="Select Site:").grid(column=0, row=1, sticky=tk.W, pady=5)
site_var = tk.StringVar()
site_dropdown = ttk.Combobox(main_frame, textvariable=site_var, state="readonly")
site_dropdown['values'] = sorted(list(TRACKER_MAPPING.keys()))
site_dropdown.grid(column=1, row=1, sticky=(tk.W, tk.E), pady=5, padx=5)

# Checkbuttons for what to send
check_frame = ttk.Frame(main_frame)
check_frame.grid(column=0, row=2, columnspan=2, sticky=tk.W, pady=5)

tracker_map_var = tk.BooleanVar(value=True)
cb_underperformance_var = tk.BooleanVar()

tracker_map_check = ttk.Checkbutton(check_frame, text="Tracker Map", variable=tracker_map_var)
tracker_map_check.pack(side=tk.LEFT, padx=5)

cb_underperformance_check = ttk.Checkbutton(check_frame, text="CB Underperformance", variable=cb_underperformance_var)
cb_underperformance_check.pack(side=tk.LEFT, padx=5)

# Send Button
send_button = ttk.Button(main_frame, text="Download and Send Email", command=download_and_send)
send_button.grid(column=0, row=3, columnspan=2, pady=10)

root.mainloop()
