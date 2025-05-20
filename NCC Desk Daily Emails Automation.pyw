#Testing Shift Summary Email Send
import pyodbc
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime, date, time
import time
from tkinter import *
from tkinter import simpledialog
import openpyxl
import pandas as pd
from tkinter import messagebox
import ctypes
import cv2
import numpy as np
import pyautogui
import pytesseract
import PIL.Image
import re
import subprocess
import sys
import os
from icecream import ic
import json
import tkinterweb


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from PythonTools import CREDS, EMAILS, PausableTimer

lily_update_file = r"G:\Shared drives\O&M\NCC Automations\Emails\Lily Updates.txt"
path_AHK = r"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
ahk_movedown_path = r"G:\Shared drives\O&M\NCC Automations\Emails\Move Lily Down.ahk"
ahk_moveback_path = r"G:\Shared drives\O&M\NCC Automations\Emails\Move Lily Back.ahk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
openEventsSheet = '1byNW58NbhuDotmdzJEFplczsXLAP0iMyS9YVU5FqQgk'
sheet_range= 'Open Events'




#X-Elio Notification Functions
def send_lily_email(solar_production, irradiance, updates):
    # Email configuration
    sender_email = EMAILS['NCC Desk']
    test = EMAILS['Joseph Lang']
    recipients = EMAILS['Lily Update List']
    print(recipients)
    smtp_port = 587
    smtp_username = EMAILS['NCC Desk']
    smtp_password = CREDS['lilyEmail']

    credentials = None
    if os.path.exists("LilyEvents-token.json"):
        credentials = Credentials.from_authorized_user_file("LilyEvents-token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(r"C:\Users\OMOPS\OneDrive - Narenco\Documents\Email Auth Creds\credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("LilyEvents-token.json", "w+") as token:
            token.write(credentials.to_json())

    try:
        service = build("sheets", "v4", credentials= credentials)
        sheets = service.spreadsheets()

        open_events = sheets.values().get(spreadsheetId=openEventsSheet, range= sheet_range).execute()
        values = open_events.get("values", [])

    except HttpError as error:
        print("Error accessing the sheet:", error)
    

    # Generate openEvents table
    openEvents_table = "<table style='border-collapse: collapse; border: 1px solid black;'>"
    for row in values:
        openEvents_table += "<tr>"
        for cell in row:
            openEvents_table += f"<td style='border: 1px solid black; padding: 8px; color: black;'>{cell}</td>"
        openEvents_table += "</tr>"
    openEvents_table += "</table>"


    # Compose email message
    msg = MIMEMultipart()
    msg['From'] = sender_email

    if test_var.get() == True:
        msg['To'] = test
    else:
        msg['To'] = ', '.join(recipients)
    ctime = datetime.now()
    print(ctime.minute)
    if 45 > ctime.minute > 15:
        hours = ctime.hour
        mins = '30'
    elif ctime.minute >= 45:
        hours = ctime.hour+1
        mins = '00'
    else:
        hours = ctime.hour
        mins = '00'
    msg['Subject'] = f'{hours}:{mins} Lily Solar Update'
    when = f'{hours}:{mins}'

    today = ctime.strftime("%m/%d/%y")

    # Email body with the provided data
    body = f'''<html>
            <body>
              <table>
                <tr>
                  <td>
                    <table border="1" style="border-collapse: collapse;">
                      <tr style="background-color: lightblue;">
                        <td><strong style="color: black;">Daily Update for Lily Solar - {today}</strong></td>
                      </tr>
                      <tr>
                        <td><strong style="color: black;">Site Performance Overview:</strong><br>
                        <span style="color: black;">As of {when}</span></td>
                      </tr>
                      <tr style="background-color: lightblue;">
                        <td><strong style="color: black;"> Current Solar Production:</strong><br>
                        <span style="color: black;">{solar_production} MW</span></td>
                      </tr>
                      <tr>
                        <td><strong style="color: black;">Irradiance:</strong><br>
                        <span style="color: black;">{irradiance} W/M²</span></td>
                      </tr>
                      <tr style="background-color: lightblue;">
                        <td><strong style="color: black;">Updates:</strong>
                          <ul>
                            <li style="color: black;">{updates}</li>
                          </ul>
                          </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td>
                    {openEvents_table} <!-- Existing table -->
                  </td>
                </tr>
              </table>
            </body>
            </html>'''

    msg.attach(MIMEText(body, 'html'))

    # Connect to Gmail SMTP server and send email
    with smtplib.SMTP("smtp.gmail.com", smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.send_message(msg)

    

    with open(lily_update_file, "w") as lilyfile:
        lilyfile.write("No New Updates")



def lily_email_data():
    os.startfile(ahk_movedown_path)
    time.sleep(2) #Wait 2 seconds for window to move.

    myconfig = r"--psm 6 --oem 3"

    try:
        findme = r"G:\Shared drives\O&M\NCC Automations\Emails\Don't Delete\to_find.png"
        location = pyautogui.locateOnScreen(findme)
        left, top, width, height = location
        new_left = int(left+102)
        loca = (new_left, int(top), int(width), int(height))
        
        sc = pyautogui.screenshot(region=loca)
        sc = cv2.cvtColor(np.array(sc), cv2.COLOR_RGB2GRAY)
        _, binary_image = cv2.threshold(sc, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        # Invert the binary image
        inverted_image = cv2.bitwise_not(binary_image)

        sc_file_path = fr"G:\Shared drives\O&M\NCC Automations\Emails\Lily Update Data\Lily_data_{time.strftime('%Y-%m-%d_%H-%M-%S')}.png"
        cv2.imwrite(sc_file_path, inverted_image)
        testy = r"G:\Shared drives\O&M\NCC Automations\Emails\Lily Update Data\Lily_data_2025-04-12_14-15-09.png"
        
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        text = pytesseract.image_to_string(inverted_image, config=myconfig)
        pattern = r"(\d+)\.?\,?(\d+)\.?(\d+)?"
        lilys = re.findall(pattern, text)
        print(lilys)
        
        irradiance = "N/A"

        for index, dig_set in enumerate(lilys, start=1):
            if len(dig_set) == 3 and dig_set[2]:
                join_digs = f"{dig_set[0]},{dig_set[1]}.{dig_set[2]}"
            else:
                join_digs = f"{dig_set[0]}.{dig_set[1]}"
            if index == 1:
                power = join_digs
            elif index == 2:
                irradiance = join_digs


        #We Should check the values of the irradiance and power and if they're way off to not send the email. Thats the simplest way to work around this and then just run the program again. Exit early time.sleep(5) then call the email function again. 
        if lilys == []:
            power = simpledialog.askfloat("Power not found", "Input MW Value (Cancel for N/A): ", parent=root)
        if power is None: #Is only None if you click Cancel
            power = "N/A"

        if irradiance == "N/A":
            irradiance = simpledialog.askfloat("Irradiance not found", "Input Irradiance Value (Cancel for N/A): ", parent=root)
        if irradiance is None:  #Is only None if you click Cancel
            irradiance = "N/A"
        os.startfile(ahk_moveback_path)

        with open(lily_update_file, "r") as lilyfile:
            update_data = lilyfile.read()

        #Delete ScreenShot
        #os.remove(sc_file_path)
        
        if power and irradiance:
            send_lily_email(power, irradiance, update_data)


    except pyautogui.ImageNotFoundException:
        def send_email():
            power = power_var.get()
            irradiance = irradiance_var.get()
            if power and irradiance:
                with open(lily_update_file, "r") as lilyfile:
                    update_data = lilyfile.read()
                send_lily_email(power, irradiance, update_data)
            img_not_fnd_win.destroy()

        img_not_fnd_win = Toplevel()
        img_not_fnd_win.title("Manual Entry Lily Email")
        img_not_fnd_win.wm_attributes("-topmost", True) 
        
        power_var = StringVar(value="00.00")
        irradiance_var = StringVar(value="N/A")

        Label(img_not_fnd_win, text="Power (MW):").grid(row=0, column=0, padx=5, pady=5)
        Entry(img_not_fnd_win, textvariable=power_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        Label(img_not_fnd_win, text="MW").grid(row=0, column=2, padx=5, pady=5)

        Label(img_not_fnd_win, text="Irradiance (W/m²):").grid(row=1, column=0, padx=5, pady=5)
        Entry(img_not_fnd_win, textvariable=irradiance_var, width=10).grid(row=1, column=1, padx=5, pady=5)
        Label(img_not_fnd_win, text="W/m²").grid(row=1, column=2, padx=5, pady=5)

        Button(img_not_fnd_win, text="Cancel", command=img_not_fnd_win.destroy).grid(row=2, column=0, padx=5, pady=5)
        Button(img_not_fnd_win, text="Send Email", command=send_email).grid(row=2, column=1, columnspan=2, padx=5, pady=5)





def lily_ask():
    response = messagebox.askyesno(parent= root, message="Are you sure?", title="Send Lily Email")
    if response == 1:
        lily_email_data()

#End of X-Elio Tasks

#Brandons Tools
def invoicing_noti():
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb;'
    connect_dbn = pyodbc.connect(db)
    c = connect_dbn.cursor()

    if datetime.today().weekday() == 4: # Friday
        table_name = "NARENCO_mf_Access_logs"
        subject_prefix = "Weekday"
    elif datetime.today().weekday() == 0: # Monday
        table_name = "NARENCO_wkEnd_Access_logs"
        subject_prefix = "Weekend"
    else:
        table_name = "NARENCO_mf_Access_logs"
        subject_prefix = "Weekday - Test"


    try:
        c.execute(f"SELECT * FROM [{table_name}]")
        entries = c.fetchall()
    except pyodbc.ProgrammingError as e:
        print(f"Error executing query for table {table_name}: {e}")
        connect_dbn.close()
        return
    
    connect_dbn.close() # Close the database connection

    if not entries:
        print(f"No entries found in {table_name} for today. Data: {entries}")
        return # No data to email
    processed_data_for_excel = []
    # Define column names based on the expected query result
    column_names = ["Unique ID", "WO No.", "Type", "Site (COUNT)",
                     "Invoice / Quote", "Description", "Last Name",
                       "Qty / Hours", "Rate", "Cost", "Markup %",
                         "Markup Total Cost", "Adddate", "Addtime",
                           "Date", "Site"]


    for entry in entries:
        start_date = entry[0]
        start_time = entry[1]
        end_date = entry[2]
        end_time = entry[3]
        location = entry[4]
        company = entry[5]
        employee = entry[6]
        peepscount = entry[7]
        notes = entry[8]

        # Initialize all 16 Excel columns, defaulting to None (which becomes blank in Excel).
        mileage_row = {col_name: None for col_name in column_names}

        mileage_row["Adddate"] = start_date.date()
        mileage_row["Date"] = start_date.date()
        mileage_row["Addtime"] = start_time.time()
        mileage_row["Type"] = 'Mileage'
        mileage_row["Site"] = location
        mileage_row["Invoice / Quote"] = 'Quote'
        mileage_row["Site (COUNT)"] = 1
        if " " in employee:
            mileage_row["Last Name"] = employee.split(" ")[-1]
        mileage_row["Description"] = notes

        processed_data_for_excel.append(mileage_row)

        # Initialize all 16 Excel columns, defaulting to None (which becomes blank in Excel).
        traveltime_row = {col_name: None for col_name in column_names}

        traveltime_row["Adddate"] = start_date.date()
        traveltime_row["Date"] = start_date.date()
        traveltime_row["Addtime"] = start_time.time()
        traveltime_row["Type"] = 'Travel Time'
        traveltime_row["Site"] = location
        traveltime_row["Invoice / Quote"] = 'Quote'
        traveltime_row["Site (COUNT)"] = 1
        if " " in employee:
            traveltime_row["Last Name"] = employee.split(" ")[-1]
        traveltime_row["Description"] = notes

        processed_data_for_excel.append(traveltime_row)

        # Initialize all 16 Excel columns, defaulting to None (which becomes blank in Excel).
        labor_row = {col_name: None for col_name in column_names}

        labor_row["Adddate"] = start_date.date()
        labor_row["Date"] = start_date.date()
        labor_row["Addtime"] = start_time.time()
        labor_row["Type"] = 'Labor'
        labor_row["Site"] = location
        labor_row["Invoice / Quote"] = 'Quote'
        labor_row["Site (COUNT)"] = 1
        if " " in employee:
            labor_row["Last Name"] = employee.split(" ")[-1]
        labor_row["Description"] = notes

                # Calculate duration for Labor row
        duration_str = "N/A"
        start_datetime_obj = None
        end_datetime_obj = None

        # Create start datetime object
        if start_date is not None and start_time is not None:
            actual_start_date = start_date.date()
            actual_start_time = start_time.time()
            start_datetime_obj = datetime.combine(actual_start_date, actual_start_time)

        # Create end datetime object
        if end_date is not None and end_time is not None:
            actual_end_date = end_date.date()
            actual_end_time = end_time.time()
            end_datetime_obj = datetime.combine(actual_end_date, actual_end_time)
        
        if start_datetime_obj and end_datetime_obj:
            if end_datetime_obj >= start_datetime_obj:
                time_difference = end_datetime_obj - start_datetime_obj
                total_seconds = time_difference.total_seconds()
                total_hours = total_seconds / 3600
                duration_str = f"{total_hours:.2f}"
        
        labor_row["Qty / Hours"] = duration_str # This now holds the calculated duration for Labor


        processed_data_for_excel.append(labor_row)

    # Create a Pandas DataFrame
    df = pd.DataFrame(processed_data_for_excel, columns=column_names)

    # Define Excel file path and name
    excel_filename = f"{subject_prefix}_Access_Logs_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    # Using a temporary directory is safer for temporary files
    excel_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp", excel_filename)
    # Ensure the directory exists (optional, temp dir usually exists)
    os.makedirs(os.path.dirname(excel_path), exist_ok=True)

    # Save DataFrame to Excel
    try:
        df.to_excel(excel_path, index=False)
        print(f"Data saved to {excel_path}")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")
        return # Don't send email if file creation failed

    # Email configuration
    sender_email = EMAILS['NCC Desk']
    recipient_email = f"{EMAILS['Joseph Lang']}, {EMAILS['Brandon Arrowood']}"
    smtp_port = 587
    smtp_username = EMAILS['NCC Desk']
    smtp_password = CREDS['shiftsumEmail']

    # Compose email message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"Invoicing Automation {subject_prefix} Access Logs - {datetime.today().strftime('%m/%d/%Y')}"

    # Add body text
    body = f"Please find the {subject_prefix} access logs attached."
    msg.attach(MIMEText(body, 'plain'))

    # Attach the Excel file
    try:
        with open(excel_path, "rb") as attachment:
            part = MIMEApplication(attachment.read(), _subtype="xlsx")
            part.add_header('Content-Disposition', 'attachment', filename=excel_filename)
            msg.attach(part)
        print(f"Attached {excel_filename}")
    except Exception as e:
        body += f"\n\nError attaching Excel file: {e}"

    # Send the email
    try:
        with smtplib.SMTP("smtp.gmail.com", smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
        print(f"Email sent successfully to {recipient_email}")
    except Exception as e:
        print(f"Error sending email: {e}")
    finally:
        # Clean up the temporary Excel file
        try:
            os.remove(excel_path)
            print(f"Removed temporary file: {excel_path}")
        except OSError as e:
            print(f"Error removing temporary file {excel_path}: {e}")

#Brandon's Functions End
    


#Shift SUMMARY Functions
def shift_Summary():
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb;'
    connect_dbn = pyodbc.connect(db)
    c = connect_dbn.cursor()
    firstentry = simpledialog.askinteger(title= "Shift Summary Query", prompt="Input the Activity Log ID from the First Entry of the Day")

    c.execute(f"SELECT * FROM [ShiftSummary] WHERE [ActivityLogID] >= {firstentry}")
    todays_entries = c.fetchall()

    todays_date = datetime.today().strftime('%m/%d/%Y')

    # Initialize a dictionary to accumulate data for each ActivityLogID
    activity_data = {}

    # Iterate over each row in the entries
    for row in todays_entries:
        activitylog_id = row[0]
        employee = row[5]
        dickey = ''.join((str(activitylog_id), str(employee)))

        # Check if the current ActivityLogID has already been processed
        if dickey not in activity_data:
            if row[8]:
                worktype = str(row[7]) if row[7] else ''
                workLog = ' | '.join((worktype, str(row[8]) if row[8] else ''))
            else:
                workLog = ''
            if row[3] and row[4] and row[5] and row[6]:
                siteAccess = ' | '.join(('', str(row[3]), str(row[4]), str(row[5]), str(row[6])))
            else:
                siteAccess = ''
            if row[13]:
                issueLog = ' | '.join((str(row[9]) if row[9] else '', str(row[10]) if row[10] else '', str(row[11]) if row[11] else '', str(row[12]) if row[12] else '', str(row[13]) if row[13] else ''))
            else:
                issueLog = ''
            
            try:
                starttime = ' '.join((str(row[14].date()), str(row[15].time())))
            except AttributeError:
                starttime = ''
                messagebox.showwarning(title="Missing Start Date/Time", message=f"Log Id: {activitylog_id}\nMissing Start Date/Time. Field will be blank in Shift Summary Table. Please input a Start Date/Time for {activitylog_id} and refresh table with 'Update Table' Button on Preview Window")
            if row[16] and row[17]:
                endtime = ' '.join((str(row[16].date()), str(row[17].time())))
            else:
                endtime = ''
            # Initialize a new entry in the dictionary
            activity_data[dickey] = {
                'location': row[1],
                'activity': row[2],
                'site_access': siteAccess,
                'site_work': [workLog],
                'issue_tracking': issueLog,
                'start_datetime': starttime,
                'end_datetime': endtime,
            }
        else:
            # Add data from row[7] and row[8] to the notes list
            if row[7] is not None and row[8] is not None:
                workLog = ' | '.join((str(row[7]), str(row[8])))
                activity_data[dickey]['site_work'].append(workLog)


    # Construct the HTML table
    html_table = f"<h2 style='text-align: center; color: black;'>NCC Daily Activity Report for {todays_date}</h2>"
    html_table += "<table style='border-collapse: collapse; width: 100%;'>"
    html_table += "<tr style='background-color: lightgray;'>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Location</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Activity</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Site Access</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Site Work</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Issue Tracking</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Start Date/Time</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>End Date/Time</th>"
    html_table += "</tr>"

    # Iterate over the accumulated data to build the HTML rows
    for dickey, data in activity_data.items():
        html_table += "<tr>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'>{data['location']}</td>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'>{data['activity']}</td>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'>{data['site_access']}</td>"
        notes_combined = "<br>".join(data['site_work'])
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'>{notes_combined}</td>"

        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'>{data['issue_tracking']}</td>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'>{data['start_datetime']}</td>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'>{data['end_datetime']}</td>"

        html_table += "</tr>"

    html_table += "</table>"
    # Display the table in a new window
    preview_window = Toplevel(root)
    preview_window.title("Shift Summary Preview")
    try:
        preview_window.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\email-1.ico")
    except Exception as e:
        print(f"Error loading icon: {e}")
    preview_window.attributes("-fullscreen", True)

    # Create a canvas and a scrollbar
    canvas = Canvas(preview_window)
    scrollbar = Scrollbar(preview_window, orient=VERTICAL, command=canvas.yview)
    scrollable_frame = Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    # Pack the scrollable_frame to expand and fill
    scrollable_frame.pack(fill=BOTH, expand=True)

    # Pack the canvas and scrollbar
    canvas.pack(fill=BOTH, expand=True)

    # Create an HtmlFrame widget to display the HTML table
    html_frame = tkinterweb.HtmlFrame(scrollable_frame)
    html_frame.load_html(html_table)
    html_frame.pack(expand=True, fill=BOTH)

    # Create buttons for restarting or sending the email
    button_frame = Frame(preview_window)
    button_frame.pack(fill=X)

    destroy_preview_button = Button(button_frame, text="Close Preview", command=preview_window.destroy, width=50)
    destroy_preview_button.pack(side=LEFT, padx=10, pady=10)

    destroy_root_button = Button(button_frame, text="Close App", command=root.destroy, width=50)
    destroy_root_button.pack(side=LEFT, padx=10, pady=10)

    restart_button = Button(button_frame, text="Update Table", command=lambda: restart_shift_summary(preview_window), width=50)
    restart_button.pack(side=RIGHT, padx=10, pady=10)

    send_button = Button(button_frame, text="Send Email", command=lambda: send_shift_Summary(html_table, preview_window), width=50)
    send_button.pack(side=RIGHT, padx=10, pady=10)

def restart_shift_summary(preview_window):
    preview_window.destroy()
    shift_Summary()

def send_shift_Summary(html_table, preview_window):
    date = datetime.now()
    today = date.strftime("%m/%d/%y")

    shift_sum_recipients = EMAILS['Shift Sum List']

    test = EMAILS['Joseph Lang']
    # Create the email message
    message = MIMEMultipart()
    message["Subject"] = f"{today} NCC Shift Summary"
    message["From"] = EMAILS['NCC Desk']
    message["To"] = ', '.join(shift_sum_recipients)
    #message["To"] = test
    password = CREDS['shiftsumEmail']
    sender = EMAILS['NCC Desk']
    
    # Attach HTML content
    html_body = MIMEText(html_table, "html")
    message.attach(html_body)

    # Send the email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(sender, shift_sum_recipients, message.as_string())    
    
    time.sleep(2.5)
    day_of_week = datetime.today().weekday()
    if day_of_week < 5: 
        os.system("shutdown /r /t 1")
    else:
        os.startfile(r"G:\Shared drives\O&M\NCC Automations\Notification System\Email Notification (Breaker).py")
        time.sleep(5)
        root.destroy()












    
myappid = 'NCC.Desk.Functions'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

root = Tk()
root.title("NCC Desk Functions")
#root.geometry("280x118")
root.wm_attributes("-topmost", True)
try:
    root.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\email-1.ico")
except Exception as e:
    print(f"Error loading icon: {e}")

cb_var = BooleanVar()
test_var = BooleanVar()

shift_summary_button = Button(root, command= lambda: shift_Summary(), text="Shift Summary", font=("Calibiri", 18), pady= 5, padx= 45)
shift_summary_button.pack(pady= 1)
lily_email_button = Button(root, command= lambda: lily_email_data(), text= "Lily Email", font=("Calibiri", 18), pady= 5, padx= 70)
lily_email_button.pack(pady= 2)
test_button = Button(root, command= lambda: invoicing_noti(), text= "Test", font=("Calibiri", 18), pady= 5, padx= 70)
test_button.pack(pady= 2)

test_frame = Frame(root)
test_frame.pack()
sched_frame = Frame(root)
sched_frame.pack()

test_CB = Checkbutton(test_frame, text= "To: Joseph Only = ✔️", variable=test_var)
test_CB.pack(side= LEFT)
auto_send_CB = Checkbutton(test_frame, text= "Lily Auto Email", variable=cb_var)
auto_send_CB.pack(side=RIGHT)

sched_label = Label(sched_frame, text="Schedule Lily Auto Send")
sched_label.grid(row=1, column= 0, columnspan= 4)

lb1var = IntVar()
ub1var = IntVar()
lb2var = IntVar()
ub2var = IntVar()
lb3var = IntVar()
ub3var = IntVar()
lb4var = IntVar()
ub4var = IntVar()

amlbl = Label(sched_frame, text="Morning Time")
amlbl.grid(row=2, column= 0, columnspan= 4)

lb1lbl = Label(sched_frame, text= "H:")
lb1lbl.grid(row=3, column=0, sticky= E)
lb1 = Entry(sched_frame, textvariable= lb1var, width= 3)
lb1.grid(row=3, column= 1, sticky= W)

lb2lbl = Label(sched_frame, text= "M:")
lb2lbl.grid(row=3, column=2, sticky= E)
lb2 = Entry(sched_frame, textvariable= lb2var, width= 3)
lb2.grid(row=3, column= 3, sticky= W)

ub1lbl = Label(sched_frame, text= "H:")
ub1lbl.grid(row=4, column=0, sticky= E)
ub1 = Entry(sched_frame, textvariable= ub1var, width= 3)
ub1.grid(row=4, column= 1, sticky= W)

ub2lbl = Label(sched_frame, text= "M:")
ub2lbl.grid(row=4, column=2, sticky= E)
ub2 = Entry(sched_frame, textvariable= ub2var, width= 3)
ub2.grid(row=4, column= 3, sticky= W)

pmlbl = Label(sched_frame, text="Afternoon Time")
pmlbl.grid(row=5, column= 0, columnspan= 4)

lb3lbl = Label(sched_frame, text= "H:")
lb3lbl.grid(row=6, column=0, sticky= E)
lb3 = Entry(sched_frame, textvariable= lb3var, width= 3)
lb3.grid(row=6, column= 1, sticky= W)

lb4lbl = Label(sched_frame, text= "M:")
lb4lbl.grid(row=6, column=2, sticky= E)
lb4 = Entry(sched_frame, textvariable= lb4var, width= 3)
lb4.grid(row=6, column= 3, sticky= W)

ub3lbl = Label(sched_frame, text= "H:")
ub3lbl.grid(row=7, column=0, sticky= E)
ub3 = Entry(sched_frame, textvariable= ub3var, width= 3)
ub3.grid(row=7, column= 1, sticky= W)

ub4lbl = Label(sched_frame, text= "M:")
ub4lbl.grid(row=7, column=2, sticky= E)
ub4 = Entry(sched_frame, textvariable= ub4var, width= 3)
ub4.grid(row=7, column= 3, sticky= W)



def check_to_send():
    check = cb_var.get()
    amhl = lb1var.get()
    amml = lb2var.get()
    amhu = ub1var.get()
    ammu = ub2var.get()

    pmhl = lb3var.get()
    pmml = lb4var.get()
    pmhu = ub3var.get()
    pmmu = ub4var.get()

    #Shift Summary Send Time
    if open_day > 4:
        target_time = datetime.now().replace(hour= 14, minute= 55, second=0, microsecond= 0)
    else:
        target_time = datetime.now().replace(hour= 18, minute= 55, second=0, microsecond= 0)
    #Morning Time
    lilyamUtarget_time = datetime.now().replace(hour= amhu, minute= ammu, second=0, microsecond= 0)
    lilyamLtarget_time = datetime.now().replace(hour= amhl, minute= amml, second=0, microsecond= 0)
    #Afternoon Time
    lilypmLtarget_time = datetime.now().replace(hour= pmhl, minute= pmml, second=0, microsecond= 0)
    lilypmUtarget_time = datetime.now().replace(hour= pmhu, minute= pmmu, second=0, microsecond= 0)
    #Now
    currenttime = datetime.now().time()


    if ((lilyamLtarget_time.time() < currenttime < lilyamUtarget_time.time()) or (lilypmLtarget_time.time() < currenttime < lilypmUtarget_time.time())) and check:
        lily_email_data()
    else:
        ic("Criteria not met", check, pmhl, pmml, lilypmLtarget_time.time(), pmhu, pmmu, lilypmUtarget_time.time())
    if currenttime >= target_time.time():
        shift_Summary()
    else: #Check Every Minute
        root.after(56000, lambda: check_to_send())



    
#Turn Autosend ON
auto_send_CB.select()
#Set AM Lily Email
lb1var.set(10)
lb2var.set(0)
ub1var.set(10)
ub2var.set(1)

#Set PM Lily Email
open_day = datetime.now().weekday()
if open_day > 4:
    lb3var.set(14)
    lb4var.set(0)
    ub3var.set(14)
    ub4var.set(1)
elif datetime.now().month in [10, 11, 12, 1, 2, 3]:
    lb3var.set(15)
    lb4var.set(0)
    ub3var.set(15)
    ub4var.set(1)
else:
    lb3var.set(17)
    lb4var.set(0)
    ub3var.set(17)
    ub4var.set(1)

check_to_send()

if datetime.now().weekday() in [0, 4]:
    invoicing_noti()

root.mainloop()
