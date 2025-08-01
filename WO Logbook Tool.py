#WO Logbook Tool
import os, re, tkinterweb, json
import pyodbc
from openpyxl import load_workbook, Workbook
import datetime
import time
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

from PythonTools import CREDS, EMAILS #Both of these Variables are Dictionaries with a single layer that holds Personnel data or app passwords

slr_customer_data = []
solt_customer_data = []
nar_customer_data = []
nce_customer_data = []
hst_customer_data = []

xelio_customer_data = []
charter_customer_data = []

location_dict = {
    1: ["*"],
    2: ["Abbot"],
    3: ["Blacktip"],
    4: ["Bluebird"],
    5: ["Cardinal"],
    6: ["Cherry Blossom"],
    7: ["Conetoe 1"],
    8: ["Cougar"],
    9: ["Duplin"],
    10: ["Freight Line"],
    11: ["Harrison"],
    12: ["Hayes"],
    13: ["Hickory"],
    14: ["Holly Swamp"],
    15: ["Lily"],
    16: ["PG Solar"],
    17: ["Tamworth"],
    18: ["Tanager"],
    19: ["Van Buren"],
    20: ["Violet"],
    21: ["Warbler"],
    22: ["Wayne I"],
    23: ["Wayne II"],
    24: ["Wayne III"],
    25: ["Wellons Farm"],
    26: ["Whitetail"],
    27: ["Whitt"],
    28: ["Charter GM"],
    29: ["Gaston PD"],
    30: ["NCC"],
    31: ["Biltmore Estates"],
    32: ["Stanly"],
    33: ["Tedder"],
    34: ["Thunderhead"],
    35: ["Volvo"],
    36: ["Marshall"],
    37: ["Shoe Show"],
    38: ["Statesboro"],
    39: ["Bulloch 1A"],
    40: ["Bulloch 1B"],
    41: ["Upson Ranch", "Upson"],
    42: ["Richmond Cadle"],
    43: ["BISHOPVILLE"],
    44: ["HICKSON"],
    45: ["JEFFERSON"],
    46: ["OGBURN"],
    47: ["Mclean"],
    48: ["Shorthorn"],
    49: ["Bluegrass"],
    50: ["Omnidian Jefferson"],
    51: ["Omnidian Stephens"],
    52: ["Harding"],
    53: ["Omnidian Davidson I"],
    54: ["BE Solar"],
    55: ["Elk"],
    56: ["Brady"],
    57: ["CDIA"],
    58: ["Sunflower"],
    59: ["Wadley"],
    60: ["Omnidian Davidson II"],
    61: ["Omnidian Sutton"],
    62: ["Omnidian Hampton"],
    63: ["Washington"],
    64: ["Gray Fox"],
    65: ["Whitehall"],
    66: ['Williams'],
    67: ['Bulloch']
}

customers_dict = {
    'slr': {'Bulloch 1A', 'Bulloch 1B', 'Elk', 'Gray Fox', 'Harding', 'Mclean', 'Richmond Cadle', 'Shorthorn', 'Sunflower', 'Upson', 'Upson Ranch', 'Warbler', 'Washington', 'Whitehall', 'Whitetail'},
    'solt': {'Conetoe 1', 'Duplin', 'Wayne I', 'Wayne II', 'Wayne III'},
    'nar': {'Bluebird', 'Cardinal', 'Cougar', 'Cherry Blossom', 'Harrison', 'Hayes', 'Hickory', 'Violet', 'Wellons Farm'},
    'nce': {'Freight Line', 'Holly Swamp', 'PG Solar'},
    'hst': {'BISHOPVILLE', 'HICKSON', 'OGBURN', 'JEFFERSON', 'Marshall', 'Tedder', 'Thunderhead', 'Van Buren'},
    'xelio': {'Lily'},
    'charter': {'Charter GM'},
}

customer_site_dict = {
    'slr': ['Bulloch 1A', 'Bulloch 1B', 'Elk', 'Gray Fox', 'Harding', 'Mclean', 'Richmond', 'Shorthorn', 'Sunflower', 'Upson', 'Upson', 'Warbler', 'Washington', 'Whitehall', 'Whitetail'],
    'solt': ['Conetoe', 'Duplin', 'Wayne I', 'Wayne II', 'Wayne III'],
    'nar': ['Bluebird', 'Cardinal', 'Cougar', 'Cherry Blossom', 'Harrison', 'Hayes', 'Hickory', 'Violet', 'Wellons'],
    'nce': ['Freight Line', 'Holly Swamp', 'PG'],
    'hst': ['Bishopville II', 'Hickson', 'Ogburn', 'Jefferson', 'Marshall', 'Tedder', 'Thunderhead', 'Van Buren'],
    'xelio': ['Lily'],
    'charter': ['Charter GM'],
}

#Tuple is as follows = name, wo-report, access_log_report 
#Will be used to identify users who get the report or do not. 
customer_report_items = [('Harrison St.', hst_customer_data, False), ('NARENCO', nar_customer_data, False), 
                         ('NCEMC', nce_customer_data, False), 
                         ('Sol River', slr_customer_data, False), ('Soltage', solt_customer_data, True)]


def dbcnxn():
    global db, connect_db, c
    #Connect to DB
    realdb = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb;'
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039 - Copy.accdb'
    if testingvar.get():
        connect_db = pyodbc.connect(db)
    else:    
        connect_db = pyodbc.connect(realdb)
    c = connect_db.cursor()

def browse_files():
    # Get the path to the Downloads folder
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    # Open file dialog to select the first file
    wo_data = filedialog.askopenfilename(
        title="Select the Emaint Report Download Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")],
        initialdir=downloads_folder
    )
    if not wo_data:
        return
    else:
        parse_wo(wo_data)




def time_validation(tvalue):
    """Validates that the input is in HH:MM format."""
    # An empty entry is a valid state
    if not tvalue:
        return True
    # The final format is 5 characters long (e.g., "23:59")
    if len(tvalue) > 5:
        return False
        
    # Check the characters as they are typed
    for i, char in enumerate(tvalue):
        if i in [0, 1, 3, 4]:  # Positions for digits
            if not char.isdigit():
                return False
        if i == 2:  # Position for the colon
            if char != ':':
                return False
                
    # Check the semantic value of the hour and minute
    if len(tvalue) >= 2:
        # Hour must be between 00 and 23
        if int(tvalue[0:2]) > 23:
            return False
    if len(tvalue) == 5:
        # Minute must be between 00 and 59
        if int(tvalue[3:5]) > 59:
            return False
            
    return True

def date_validation(dvalue):
    """Validates that the input is in mm/dd/yyyy format."""
    # An empty entry is a valid state
    if not dvalue:
        return True
    # The final format is 10 characters long
    if len(dvalue) > 10:
        return False

    # Check the characters as they are typed
    for i, char in enumerate(dvalue):
        if i in [0, 1, 3, 4, 6, 7, 8, 9]:  # Positions for digits
            if not char.isdigit():
                return False
        if i in [2, 5]:  # Positions for slashes
            if char != '/':
                return False

    # Check the semantic value of the month and day
    if len(dvalue) >= 2:
        # Month must be between 01 and 12
        month = int(dvalue[0:2])
        if month < 1 or month > 12:
            return False
    if len(dvalue) >= 5:
        # Day must be between 01 and 31
        day = int(dvalue[3:5])
        if day < 1 or day > 31:
            return False
    
    return True
    
def parse_wo(wos):
    global c
    dbcnxn()

    activity_logbook_data = []
    wb = load_workbook(wos)
    ws = wb.active
    for index, row in enumerate(ws.iter_rows(values_only=True)):
        if index == 0:
            continue
        wo = None
        startdate = None
        starttime = None
        enddate = None
        endtime = None
        site = None
        activityid = None
        locationid = None
        userlistid = None
        issuelistid = None
        editDate = None
        activity_log_id = None
        wo_notes = None
        customer_note = None

        # Assuming row[10] contains HTML content
        clean_text = BeautifulSoup(row[8], "html.parser").get_text()

        if row[0] == "JOSEPHL99":
            userlistid = 7
        elif row[0] == "JACOBBUD99":
            userlistid = 10
        elif row[0] in {"NARENCO", "JAYMEORR99", "BRANDONA99", "BRANDONA96","NEWMANSE99"}:
            userlistid = 3
        else:
            messagebox.showerror(title="Exiting WO Input", message=f"Found Error in WO XL File\nUser is Neither:\nJoseph Lang\nJacob Budd\nBrandon Arrowood (Admin Account)\nWO: {wo}")
            root.destroy()

        wo = row[2]
        site = row[3]
        activityid = 3

        locationid = None
        for key, locations in location_dict.items():
            if site in locations:
                locationid = key
                break
        if locationid is None:
            messagebox.showerror(title="Site Not Found", message=f"Site '{site}' not found in location dictionary.\nWO: {wo}")
            root.destroy()

        for customervar, siteList in customers_dict.items():
            if site in siteList:
                customer = customervar
                break

        startdate_match = re.search(r'Start Date: (\d{1,2}[-/]\d{1,2}(?:[-/]\d{2,4})?)', clean_text)
        starttime_match = re.search(r'Start Time: (\d{1,2}[:;]?\d{1,2}|\d{3,4})', clean_text)
        enddate_match = re.search(r'End Date: (\d{1,2}[-/]\d{1,2}(?:[-/]\d{2,4})?)', clean_text)
        endtime_match = re.search(r'End Time: (\d{1,2}[:;]?\d{1,2}|\d{3,4})', clean_text)
        # Function to add current year if missing
        def add_current_year(date_str):
            # Check if the year is missing
            if re.match(r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', date_str):
                return date_str  # Year is present
            else:
                current_year = datetime.datetime.now().year
                return f"{date_str}/{current_year}"
        def format_time(time_str):
            # If the time is already in HH:MM format, return it as is
            if ':' in time_str:
                return time_str
            # If the time is a group of 3 or 4 digits, format it to HH:MM
            elif len(time_str) == 3:
                return f"0{time_str[0]}:{time_str[1:]}"
            elif len(time_str) == 4:
                return f"{time_str[:2]}:{time_str[2:]}"
            else:
                raise ValueError("Invalid time format")

        # Extract and format the start time
        if starttime_match:
            #print(starttime_match.group(1), clean_text)
            starttime = format_time(starttime_match.group(1))
            #print("Formatted Start Time:", starttime)

        # Extract and format the end time
        if endtime_match:
            endtime = format_time(endtime_match.group(1))
            #print("Formatted End Time:", endtime)

        # Extract the end date
        if enddate_match:
            enddate = add_current_year(enddate_match.group(1))
        # Extract the end date
        if startdate_match:
            startdate = add_current_year(startdate_match.group(1))
        editDate = datetime.date.today()
        
        # Insert data into the ActivityLog table and retrieve the ActivityLogID
        try:
            
            # Execute the insert statement
            c.execute("""
                INSERT INTO ActivityLog (ActivityListID, UserListID, LocationListID, StartDate, StartTime, EndDate, EndTime, EditDate)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (activityid, userlistid, locationid, startdate, starttime, enddate, endtime, editDate))
            
            # Retrieve the last inserted ActivityLogID
            c.execute("SELECT @@IDENTITY")
            activity_log_id = c.fetchone()[0]

            connect_db.commit()
            

        except Exception as e:
            messagebox.showerror(title="Database Error", message=f"An error occurred while inserting data into the database: {e}")
        
        stowmatch1 = re.search(r'stow', row[6], re.IGNORECASE)
        stowmatch2 = re.search(r'stow', row[7], re.IGNORECASE)
        stowmatch3 = re.search(r'stow', row[8], re.IGNORECASE)


        invnomatch = re.search(r'Inverter\s*(\d+)', row[6], re.IGNORECASE)
        if invnomatch:
            invno = invnomatch.group(1)
        else:
            invno = '0'

        invcheck1 = re.search(r'(Inverter|INV)', row[6], re.IGNORECASE)
        invcheck2 = re.search(r'(Inverter|INV)', row[7], re.IGNORECASE)

        underperformcheck1 = re.search(r'underperform', clean_text, re.IGNORECASE)
        underperformcheck2 = re.search(r'underperform', row[7], re.IGNORECASE)

        trcheck1 = re.search(r'(Node|TCU|Tracker)', row[6], re.IGNORECASE)
        trcheck2 = re.search(r'(Node|TCU|Tracker)', row[7], re.IGNORECASE)

        cbcheck1 = re.search(r'(CMB|CB|Combiner Box)', row[6], re.IGNORECASE)
        cbcheck2 = re.search(r'(CMB|CB|Combiner Box)', row[7], re.IGNORECASE)

        #Remote Repair Check
        rrcheck = re.search(r'Can we resolve the issue Remotely\? \(Y or N\) (Y|Yes|N|No|None)', clean_text, re.IGNORECASE)
        #print(rrcheck, clean_text)
        if rrcheck:
            response = rrcheck.group(1).lower()
            #print(response)
            if response in ['none', 'n', 'no']:
                repairs_note = "by the technician on site"
            else:
                repairs_note = 'by the NCC remotely'



        if row[4] == 'Site Outage' and (row[5] == 'Utility Trip' or row[5] == 'Utility Planned'):
            issuelistid = 165
            if enddate is None:
                customer_note = f'Utility Tripped the site offline at {starttime} on {startdate}'
            else:
                customer_note = f'Utility Tripped the site offline at {starttime} on {startdate}. Production restored to the site {repairs_note} at {endtime} on {enddate}'
        elif row[4] == 'Site Outage' and (row[5] == 'Site Trip' or row[5] == 'Nuisance Trip'):
            issuelistid = 167
            if enddate is None:
                customer_note = f'Site Tripped offline at {starttime} on {startdate}'
            else:
                customer_note = f'Site Tripped offline at {starttime} on {startdate}. Production restored to the site {repairs_note} at {endtime} on {enddate}'
        elif row[4] == 'Equipment Outage' and (invcheck1 is not None or invcheck2 is not None):
            issuelistid = 42
            if enddate is None:
                customer_note = f'Inverter {invno} tripped offline at {starttime} on {startdate}'
            else:
                customer_note = f'Inverter {invno} tripped offline at {starttime} on {startdate}. Production restored to the device {repairs_note} at {endtime} on {enddate}'
        elif row[4] == 'COMMs Outage' and (invcheck1 is not None or invcheck2 is not None):
            issuelistid = 64
        elif row[4] == 'Underperformance' and (underperformcheck1 is not None or underperformcheck2 is not None) and (invcheck1 is not None or invcheck2 is not None):
            issuelistid = 60
        elif row[4] == 'Underperformance' and (trcheck1 is not None or trcheck2 is not None):
            issuelistid = 229
        elif row[4] == 'Underperformance' and (cbcheck1 is not None or cbcheck2 is not None):
            issuelistid = 212
        elif stowmatch1 or stowmatch2 or stowmatch3:
            issuelistid = 227
            customer_note = 'Stowed for Local Thunderstorm or Excessive Wind Speeds'
        else:
            issuelistid = None

        wo_notes_match = re.search(r'Summary of issue:\s*(.*?)\s*Can we resolve the issue Remotely\?', clean_text, re.DOTALL)
        if wo_notes_match:
            wo_notes = f'WO {wo}  {wo_notes_match.group(1)}'
        else:
            wo_notes = f'WO {wo}'
        
        try:
            # Now you can use activity_log_id to insert more data into another table
            # Example: Insert into another table using activity_log_id
            
            c.execute("""
                INSERT INTO ObservationLog (ActivityLogID, DeviceListID, IssueListID, Notes)
                VALUES (?, ?, ?, ?)
            """, (activity_log_id, None, issuelistid, wo_notes))
            connect_db.commit()
            
        except Exception as e:
            messagebox.showerror(title="Database Error", message=f"An error occurred while inserting data into the database: {e}")
            root.destroy()
        

        #Customer Notification Process
        if customer_note is not None:
            globals()[f'{customer}_customer_data'].append([wo, site, startdate, starttime, enddate, endtime, customer_note]) 

    if messagebox.askokcancel(title="WO's to Logbook Success!",
    message= "Successfully input all WO's to the NCC Logbook\nClick OK to view Customer Reports\nClick Cancel to Close"):
        for customer, customer_data, site_access_log in customer_report_items:
            if site_access_log:
                site_access_table = site_access_query(customer)
                customer_noti(customer, customer_data, site_access_table)
            else:
                customer_noti(customer, customer_data, site_access_log)
    else:
        connect_db.close()
        root.destroy()

def send_email(customer_data, window, customer):
    #print(f"Custom Data: {customer_data}")
    # Create the email message
    today = datetime.date.today().strftime("%m/%d/%y")
    message = MIMEMultipart()
    me = EMAILS['Joseph Lang']
    brandon = EMAILS['Brandon Arrowood']
    if customer == 'Harrison St.':
        message["Subject"] = f"NARENCO O&M | {today} Incident Report - {customer}"
        recipients = EMAILS['Harrison St']
    elif customer == 'Sol River':
        message["Subject"] = f"NARENCO O&M | {today} Incident Report - {customer}"
        recipients = EMAILS['Sol River']
    elif customer == 'Soltage':
        message["Subject"] = f"NARENCO O&M | {today} Daily Report - {customer}"
        site_access = site_access_query('Soltage')
        message.attach(MIMEText(site_access, "html"))
        recipients = EMAILS['Soltage']
    elif customer == 'NCEMC':
        message["Subject"] = f"NARENCO O&M | {today} Incident Report - {customer}"
        recipients = EMAILS['NCEMC']
    elif customer == 'NARENCO':
        message["Subject"] = f"NARENCO O&M | {today} Incident Report - {customer}"
        recipients = EMAILS['NARENCO']
    

    message["From"] = EMAILS['NCC Desk']
    if testingvar.get():
        message["To"] = me
    else:
        message["To"] = ', '.join(recipients)

    password = CREDS['shiftsumEmail']
    sender = EMAILS['NCC Desk']

    # Title/Header
    email_table = f"<h2 style='text-align: center; color: black;'>NCC Daily WO Report - {customer}</h2>"
    # Initialize HTML table with border style and header row
    email_table += "<table style='border-collapse: collapse; width: 100%;'>"
    # Header row with lightgray background
    email_table += "<tr style='background-color: lightgray;'>"
    email_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Ticket #</th>"
    email_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Site</th>"
    email_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Start Date | Time</th>"
    email_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>End Date | Time</th>"
    email_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Issue</th>"
    email_table += "</tr>"
    
    # Iterate over each row in the entries
    for row in customer_data:
        # Start a new row
        email_table += "<tr>"
        email_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{row[0]}</td>"
        email_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{row[1]}</td>"
        email_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{row[2]} | {row[3]}</td>"
        if row[4] != 'None':
            email_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{row[4]} | {row[5]}</td>"
        else:
            email_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'> </td>"
        email_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'>{row[6]}</td>"
        # End the row
        email_table += "</tr>"
    # End the table
    email_table += "</table>"
    if customer_data:
        message.attach(MIMEText(email_table, "html"))

    # Send the email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        if testingvar.get():
            server.sendmail(sender, me, message.as_string())
        else:
            server.sendmail(sender, recipients, message.as_string())
    window.destroy()

def site_access_query(customer):
    global c
    dbcnxn()
    if customer == 'Soltage':
        sites = [7, 9, 22, 23, 24]
    #Add Other customers LocationIDs according to the NCC 039 Database
    #elif customer == '?'

    query = """
    SELECT al.StartDate, al.StartTime, al.EndDate, al.EndTime, a.PeopleCount, cl.Company, l.Location
    FROM ((AccessLog AS a
    INNER JOIN ActivityLog  AS al ON a.ActivityLogID = al.ActivityLogID)
    INNER JOIN ContactList AS cl ON a.ContactListID = cl.ContactListID)
    INNER JOIN LocationList as l ON al.LocationListID = l.LocationID
    WHERE al.LocationListID IN ({})
    AND al.StartDate = ?
    """.format(','.join('?' for _ in sites))
    
    today = datetime.date.today()  
    
    c.execute(query, *sites, today)
    activity_log_results = c.fetchall()

        # Title/Header
    html_table = "<h2 style='text-align: center; color: black;'>Access Log Report</h2>"
    # Initialize HTML table with border style and header row
    html_table += "<table style='border-collapse: collapse; width: 100%;'>"
    # Header row with lightgray background
    html_table += "<tr style='background-color: lightgray;'>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Location</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Company</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'># Personnel</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Start Time</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>End Time</th>"
    html_table += "</tr>"
    # Iterate over each row in the results
        # Combine rows with identical data except for PeopleCount
    combined_results = {}
    for row in activity_log_results:
        key = (row.StartDate, row.StartTime, row.EndDate, row.EndTime, row.Company, row.Location)
        if key in combined_results:
            combined_results[key] += row.PeopleCount
        else:
            combined_results[key] = row.PeopleCount

    for key, people_count in combined_results.items():
        start_date, start_time, end_date, end_time, company, location = key
        # Start a new row
        html_table += "<tr>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{location}</td>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{company}</td>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{people_count}</td>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{start_time.strftime("%H:%M")}</td>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{end_time.strftime("%H:%M") if end_time else ""}</td>"
        # End the row
        html_table += "</tr>"
    
    if activity_log_results == []:
        html_table += "<tr>"
        html_table += f"<td style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;' colspan= '5'>No Site Access Today</td>"
        html_table += "</tr>"

    # End the table
    html_table += "</table>"
    connect_db.close()
    return html_table

def site_access_only():
    for customer, table, site_accessTF in customer_report_items:
        if site_accessTF:
            site_access_table = site_access_query(customer)
            customer_noti(customer, table, site_access_table)

def manual_wo_reports():
    for customer, customer_data, site_access_log in customer_report_items:
        if site_access_log:
            site_access_table = site_access_query(customer)
            customer_noti(customer, customer_data, site_access_table)
        else:
            customer_noti(customer, customer_data, site_access_log)

def customer_noti(customer, customer_data, site_access_table):
    # Title/Header
    html_table = f"<h2 style='text-align: center; color: black;'>NCC Daily WO Report - {customer}</h2>"
    # Initialize HTML table with border style and header row
    html_table += "<table style='border-collapse: collapse; width: 100%;'>"
    # Header row with lightgray background
    html_table += "<tr style='background-color: lightgray;'>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Ticket #</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Site</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Start Date | Time</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>End Date | Time</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Issue</th>"
    html_table += "</tr>"
    
    # Iterate over each row in the entries
    for row in customer_data:
        # Start a new row
        html_table += "<tr>"
        html_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{row[0]}</td>"
        html_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{row[1]}</td>"
        html_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{row[2]} | {row[3]}</td>"
        if row[4]:
            html_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue; text-align: center;'>{row[4]} | {row[5]}</td>"
        else:
            html_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'> </td>"
        html_table += f"<td contenteditable='true' style='border: 1px solid black; padding: 8px; color: black; background-color: lightblue;'>{row[6]}</td>"
        # End the row
        html_table += "</tr>"
    # End the table
    html_table += "</table>"
    # Display the table in a new window
    preview_window = tk.Toplevel(root)
    preview_window.title(f"{customer} Report Preview")
    try:
        # Note: The icon path should be valid for this to work.
        # preview_window.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\email-1.ico")
        pass # Pass for now to make it runnable
    except Exception as e:
        print(f"Error loading icon: {e}")
    preview_window.geometry("1000x800")
    preview_window.resizable(True, True)

    # Create a container frame for the canvas and scrollbar
    container_frame = tk.Frame(preview_window)
    container_frame.pack(fill="both", expand=True)

    # Create a canvas and a scrollbar
    canvas = tk.Canvas(container_frame)
    scrollbar = tk.Scrollbar(container_frame, orient="vertical", command=canvas.yview)
    
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollable_frame = tk.Frame(canvas)
    scrollable_frame_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    def on_frame_configure(event=None):
        """Reset the scroll region to encompass the inner frame"""
        canvas.configure(scrollregion=canvas.bbox("all"))

    scrollable_frame.bind("<Configure>", on_frame_configure)
    canvas.bind(
        "<Configure>",
        lambda e: FrameWidth(e, canvas, scrollable_frame_id)
    )
    
    def on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", on_mousewheel)


    if site_access_table:
        site_frame = tkinterweb.HtmlFrame(scrollable_frame)
        site_frame.load_html(site_access_table)
        site_frame.pack(expand=True, fill="x")

    html_frame = tkinterweb.HtmlFrame(scrollable_frame)
    html_frame.load_html(html_table)
    html_frame.pack(expand=True, fill="x")

    #Instructions
    main_lbl = tk.Label(scrollable_frame, text='Edit the Table above with the one Below', borderwidth=1, relief="solid", font=('TkDefaultFont', 16))
    main_lbl.pack(expand=True, fill='x')
    
    # Create a frame for the grid-based table
    grid_frame = tk.Frame(scrollable_frame)
    grid_frame.pack(expand=True, fill='both')

    # Add headers to the grid
    headers = ["✓= Include", "WO", "Site", "Start Date", "Start Time", "End Date", "End Time", "Issue"]
    for col, header in enumerate(headers):
        label = tk.Label(grid_frame, text=header, borderwidth=1, relief="solid", padx=5, pady=5)
        label.grid(row=0, column=col, sticky="nsew")

    entry_widgets = []  # List to store entry widgets
    check_vars = []  # List to store checkbox variables

    # Populate the grid with existing data
    for row_index, row_data in enumerate(customer_data, start=1):
        entry_row = []
        check_var = tk.BooleanVar(value=True)
        check_button = tk.Checkbutton(grid_frame, variable=check_var)
        check_button.grid(row=row_index, column=0, sticky="nsew")
        check_vars.append(check_var)

        # The original data has 7 columns, so we map them to the 7 entry columns
        for col_index, value in enumerate(row_data):
            if col_index in [0, 1, 6]:
                entry = tk.Entry(grid_frame, borderwidth=1, relief="solid")
                entry.insert(0, str(value))
                entry.grid(row=row_index, column=col_index + 1, sticky="nsew")
                entry_row.append(entry)
            elif col_index in [2, 4]:
                entry = tk.Entry(grid_frame, borderwidth=1, relief="solid", validate='all', validatecommand=vcmd_date)
                entry.insert(0, str(value))
                entry.grid(row=row_index, column=col_index + 1, sticky="nsew")
                entry_row.append(entry)
            elif col_index in [3, 5]:
                entry = tk.Entry(grid_frame, borderwidth=1, relief="solid", validate='all', validatecommand=vcmd_time)
                entry.insert(0, str(value))
                entry.grid(row=row_index, column=col_index + 1, sticky="nsew")
                entry_row.append(entry)
        entry_widgets.append(entry_row)

    # Configure the grid columns to be responsive
    for i in range(7):
        grid_frame.grid_columnconfigure(i, weight=1, minsize=30)
    grid_frame.grid_columnconfigure(7, weight=10, minsize=100)


    def add_new_row(customer):
        """Creates a new empty row with a checkbox and entry widgets."""
        # Calculate the next available row index
        row_index = len(entry_widgets) + 1
        
        entry_row = []
        check_var = tk.BooleanVar(value=True)
        check_button = tk.Checkbutton(grid_frame, variable=check_var)
        check_button.grid(row=row_index, column=0, sticky="nsew")
        check_vars.append(check_var)

        # Create 7 empty entry widgets for the new row
        for col_index in range(7):
            print(col_index)
            if col_index in [0, 6]:
                entry = tk.Entry(grid_frame, borderwidth=1, relief="solid")
                entry.grid(row=row_index, column=col_index + 1, sticky="nsew")
                entry_row.append(entry)
            elif col_index in [2, 4]:
                entry = tk.Entry(grid_frame, borderwidth=1, relief="solid", validate='all', validatecommand=vcmd_date)
                entry.grid(row=row_index, column=col_index + 1, sticky="nsew")
                entry_row.append(entry)
            elif col_index in [3, 5]:
                entry = tk.Entry(grid_frame, borderwidth=1, relief="solid", validate='all', validatecommand=vcmd_time)
                entry.grid(row=row_index, column=col_index + 1, sticky="nsew")
                entry_row.append(entry)
            elif col_index == 1:
                site = tk.StringVar()
                sitedd = ttk.Combobox(grid_frame, textvariable= site)
                if customer == "Harrison St.":
                    sitedd['values'] = customer_site_dict['hst']
                elif customer == "NARENCO":
                    sitedd['values'] = customer_site_dict['nar']
                elif customer == "NCEMC":
                    sitedd['values'] = customer_site_dict['nce']
                elif customer == "Soltage":
                    sitedd['values'] = customer_site_dict['solt']
                elif customer == "Sol River":
                    sitedd['values'] = customer_site_dict['slr']
                sitedd.grid(row=row_index, column=col_index + 1, sticky="nsew")
                entry_row.append(sitedd)
        
        entry_widgets.append(entry_row)
    add_new_row(customer)
    custom_data = []
    def collect_data():
        custom_data.clear()
        for entry_row, check_var in zip(entry_widgets, check_vars):
            if check_var.get():
                row_data = [entry.get() for entry in entry_row]
                custom_data.append(row_data)

    button_frame = tk.Frame(preview_window)
    button_frame.pack(fill="x", side="bottom")

    send_button = tk.Button(button_frame, text="Send Email -> Next Preview", command=lambda: [collect_data(), send_email(custom_data, preview_window, customer)], bg='lightgreen')
    send_button.pack(side="left", fill='x', expand=True)

    add_row_button = tk.Button(button_frame, text="Add Row", command=lambda: add_new_row(customer), bg='lightblue')
    add_row_button.pack(side="left", fill='x', expand=True)

    next_preview_button = tk.Button(button_frame, text="Next Preview", command=preview_window.destroy, bg='yellow')
    next_preview_button.pack(side="left", fill='x', expand=True)

    close_button = tk.Button(button_frame, text="Close Program", command=root.destroy, bg='orange')
    close_button.pack(side="left", fill='x', expand=True)


def FrameWidth(event, canvas, scrollable_frame):
    # Update the scrollregion of the canvas to match the size of the scrollable frame
    canvas.configure(scrollregion=canvas.bbox("all"))
    # Adjust the width of the scrollable frame to match the container frame's width
    canvas.itemconfig(scrollable_frame, width=event.width)

    
# Set up the tkinter window
root = tk.Tk()
root.title("WO Processing Tool")
vcmd_date = (root.register(date_validation), '%P')
vcmd_time = (root.register(time_validation), '%P')
testingvar = tk.BooleanVar()
testingCB = tk.Checkbutton(root, text="Testing Mode", variable=testingvar)
testingCB.pack()

# Create a button to browse files
browse_button = tk.Button(root, text="Select Daily WO File", command=browse_files, height=3, width=40)
browse_button.pack()
browse_button = tk.Button(root, text="Site Access Reports Only", command=site_access_only, height=3, width=40)
browse_button.pack()
browse_button = tk.Button(root, text="Manual Reports\n(No Emaint Excel File)", command=manual_wo_reports, height=3, width=40)
browse_button.pack()
# Run the tkinter main loop
root.mainloop()