import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys, os
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Google API Imports
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
import io

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import EMAILS, CREDS, get_google_credentials

sites_spd_map = {
    "Bulloch 1A": "1wZsj_yo4QA3hLkrocVvDfYd4U6W4OvYUegl_FhMqQdQ",
    "Bulloch 1B": "11cx6M5rrmlY_Jvrky--Mtbo-U-De7W7h8qRWoMQQfJU",
    "Richmond": "1ulZYEkcOMeK1urDrN3cXqUVa8ISa7bhRXJdT5zRXTQg",
    "Upson": "1cFFVEm74ugTT3WGQUMfRkLMj0zx-D5LuymGJfS_GVnE",
    "Whitehall": "1uQ4Wc-o6-HioUU0Tm4R3c73maYw_CVItQGroENuLNvY",
}


def send_email(recipient_email, site_name, attachment_path):
    """Sends an email with the tracker map as an Excel attachment."""
    sender_email = EMAILS['NCC Desk']
    smtp_password = CREDS['shiftsumEmail']
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    date_str = datetime.now().strftime("%Y-%m-%d")

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ', '.join(recipient_email) if isinstance(recipient_email, list) else recipient_email
    msg['Subject'] = f"Tracker Data {date_str} for {site_name}"

    # HTML body
    body = f"""
    <html>
        <body>
            <p>Hello,</p>
            <p>Attached is the tracker data for {site_name} for {date_str}.</p>
            <p>This is an accurate example of the final output, sent to technicians scheduled for tracker repairs.<br>
            All feed back is welcomed!</p>

            <p>Thank you,<br>NCC Automation</p>
        </body>
    </html>
    """
    msg.attach(MIMEText(body, 'html'))

    # Attach the file
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

    if not technician or not site:
        messagebox.showwarning("Selection Missing", "Please select both a technician and a site.")
        return

    recipient_email = EMAILS.get(technician)
    spreadsheet_id = sites_spd_map.get(site)

    if not recipient_email or not spreadsheet_id:
        messagebox.showerror("Error", "Could not find email or spreadsheet ID for your selection.")
        return

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
    date_str = datetime.now().strftime("%Y-%m-%d")
    file_path = os.path.join(os.path.expanduser("~"), "Downloads", f"{site}_tracker_map_{date_str}.xlsx")

    try:
        creds = get_google_credentials()
        drive_service = build('drive', 'v3', credentials=creds)

        request = drive_service.files().export_media(fileId=spreadsheet_id,
                                                     mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"Download {int(status.progress() * 100)}%.")

        # Save the downloaded content to a file
        with open(file_path, 'wb') as f:
            f.write(fh.getvalue())

        # Send the email
        send_email(recipient_email, site, file_path)

    except HttpError as error:
        messagebox.showerror("API Error", f"An error occurred with the Google Drive API: {error}")

# --- GUI Setup ---
root = tk.Tk()
root.title("Technician Tracker Map Sender")

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
site_dropdown['values'] = sorted(list(sites_spd_map.keys()))
site_dropdown.grid(column=1, row=1, sticky=(tk.W, tk.E), pady=5, padx=5)

# Send Button
send_button = ttk.Button(main_frame, text="Download and Send Email", command=download_and_send)
send_button.grid(column=0, row=2, columnspan=2, pady=10)

root.mainloop()
