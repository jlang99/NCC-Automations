#WO Logbook Tool
import os, re, tkinterweb, json
import pyodbc
from openpyxl import load_workbook, Workbook
import datetime
import time
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog, messagebox

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

with open(r"G:\Shared drives\O&M\NCC Automations\Credentials\app credentials.json", 'r') as credsfile:
    creds = json.load(credsfile)

slr_customer_data = []
solt_customer_data = []
nar_customer_data = []
nce_customer_data = []
hst_customer_data = []

xelio_customer_data = []

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
    65: ["Whitehall"]
}

customers_dict = {
    'slr': ['Bulloch 1A', 'Bulloch 1B', 'Elk', 'GRay Fox', 'Harding', 'Mclean', 'Richmond Cadle', 'Shorthorn', 'Sunflower', 'Upson', 'Upson Ranch', 'Warbler', 'Washington', 'Whitehall', 'Whitetail'],
    'solt': ['Conetoe 1', 'Duplin', 'Wayne I', 'Wayne II', 'Wayne III'],
    'nar': ['Bluebird', 'Cardinal', 'Cougar', 'Cherry Blossom', 'Harrison', 'Hayes', 'Hickory', 'Violet', 'Wellons'],
    'nce': ['Freight Line', 'Holly Swamp', 'PG Solar'],
    'hst': ['BISHOPVILLE', 'HICKSON', 'OGBURN', 'JEFFERSON', 'Marshall', 'Tedder', 'Thunderhead', 'Van Buren'],
    'xelio': ['Lily']
}





def dbcnxn():
    global db, connect_db, c
    #Connect to DB
    realdb = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb;'
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039 - Copy.accdb'
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
        else:
            messagebox.showerror(title="Exiting WO Input", message=f"Found Error in WO XL File\nUser is Neither:\nJoseph Lang\nJacob Budd\nWO: {wo}")
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
            print(starttime_match.group(1), clean_text)
            starttime = format_time(starttime_match.group(1))
            print("Formatted Start Time:", starttime)

        # Extract and format the end time
        if endtime_match:
            endtime = format_time(endtime_match.group(1))
            print("Formatted End Time:", endtime)

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
        customer_noti(hst_customer_data, "Harrison St.")
        customer_noti(slr_customer_data, "Sol River")
        customer_noti(solt_customer_data, "Soltage")
        customer_noti(nce_customer_data, "NCMEC")
        customer_noti(nar_customer_data, "NARENCO")
    else:
        root.destroy()

def send_email(customer_data, window, customer):
    #print(f"Custom Data: {customer_data}")
    me = 'joseph.lang@narenco.com'
    if customer == 'Harrison St.':
        recipients = ['jayme.orrock@narenco.com', 'cdepodesta@harrisonst.com', 'brandon.arrowood@narenco.com', 'jhopkins@harrisonst.com', 'HS_DG_Solar@harrisonst.com', 'cclark@harrisonst.com', 'cgao@harrisonst.com', 'newman.segars@narenco.com']
    elif customer == 'Sol River':
        recipients = ['brandon@solrivercapital.com', 'projects@solrivercapital.com', 'jayme.orrock@narenco.com', 'brandon.arrowood@narenco.com', 'newman.segars@narenco.com']
    elif customer == 'Soltage':
        recipients = ['assetmanagement@soltage.com', 'blamorticella@soltage.com', 'rgray@soltage.com', 'hgao@soltage.com', 'operations@soltage.com', 'jayme.orrock@narenco.com', 'brandon.arrowood@narenco.com', 'newman.segars@narenco.com']
    elif customer == 'NCEMC':
        recipients = ['jayme.orrock@narenco.com', 'brandon.arrowood@narenco.com', 'newman.segars@narenco.com', 'amy.roswick@ncemcs.com', 'john.cook@ncemcs.com', 'Carlton.Lewis@ncemcs.com']
    elif customer == 'NARENCO':
        recipients = ['jayme.orrock@narenco.com', 'brandon.arrowood@narenco.com', 'newman.segars@narenco.com', 'andrew.giraldo@narenco.com', 'mark.caddell@narenco.com', 'jesse.montgomery@cleanshift.energy']
    
    today = datetime.date.today().strftime("%m/%d/%y")
    # Create the email message
    message = MIMEMultipart()
    message["Subject"] = f"NARENCO O&M | {today} Incident Report - {customer}"
    message["From"] = "omops@narenco.com"
    #message["To"] = me
    message["To"] = ', '.join(recipients)

    password = creds['credentials']['shiftsumEmail']
    sender = "omops@narenco.com"

    # Title/Header
    email_table = f"<h2 style='text-align: center; color: black;'>NCC Daily WO Report - {customer}</h2>"
    # Initialize HTML table with border style and header row
    email_table += "<table style='border-collapse: collapse; width: 100%;'>"
    # Header row with lightgray background
    email_table += "<tr style='background-color: lightgray;'>"
    email_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Ticket #</th>"
    email_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Site</th>"
    email_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Start Date/Time</th>"
    email_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>End Date/Time</th>"
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

    # Attach the edited HTML content
    html_body = MIMEText(email_table, "html")
    message.attach(html_body)

    # Send the email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(sender, me, message.as_string())
    window.destroy()

def customer_noti(customer_data, customer):
    if customer_data == []:
        return
    # Title/Header
    html_table = f"<h2 style='text-align: center; color: black;'>NCC Daily WO Report - {customer}</h2>"
    # Initialize HTML table with border style and header row
    html_table += "<table style='border-collapse: collapse; width: 100%;'>"
    # Header row with lightgray background
    html_table += "<tr style='background-color: lightgray;'>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Ticket #</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Site</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>Start Date/Time</th>"
    html_table += "<th style='border: 1px solid black; padding: 8px; color: black;'>End Date/Time</th>"
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
    preview_window.title("Shift Summary Preview")
    try:
        preview_window.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\email-1.ico")
    except Exception as e:
        print(f"Error loading icon: {e}")
    preview_window.attributes("-fullscreen", True)

    # Create a canvas and a scrollbar
    canvas = tk.Canvas(preview_window)
    scrollbar = tk.Scrollbar(preview_window, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    # Pack the scrollable_frame to expand and fill
    scrollable_frame.pack(fill="both", expand=True)

    # Pack the canvas and scrollbar
    canvas.pack(fill="both", expand=True)

    
    # Create an HtmlFrame widget to display the HTML table
    html_frame = tkinterweb.HtmlFrame(scrollable_frame)
    html_frame.load_html(html_table)
    html_frame.pack(expand=True, fill="both")
    
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

    for row_index, row in enumerate(customer_data, start=1):
        entry_row = []  # List to store entries for this row
        check_var = tk.BooleanVar(value=True)  # Checkbox variable, default to True
        check_button = tk.Checkbutton(grid_frame, variable=check_var)
        check_button.grid(row=row_index, column=0, sticky="nsew")
        check_vars.append(check_var)

        for col_index, value in enumerate(row):
            entry = tk.Entry(grid_frame, borderwidth=1, relief="solid")
            entry.insert(0, str(value))
            entry.grid(row=row_index, column=col_index + 1, sticky="nsew")
            entry_row.append(entry)
        entry_widgets.append(entry_row)

        # Configure the grid columns
        for i in range(7):
            grid_frame.grid_columnconfigure(i, weight=1, minsize=30)  # Smaller columns

        # Configure the last column to take the remaining space
        grid_frame.grid_columnconfigure(7, weight=100, minsize=100) 

    custom_data = []
    def collect_data():
        custom_data.clear()  # Clear existing data
        for entry_row, check_var in zip(entry_widgets, check_vars):
            if check_var.get():  # Only include rows where the checkbox is selected
                row_data = [entry.get() for entry in entry_row]
                custom_data.append(row_data)

    button_frame = tk.Frame(preview_window)
    button_frame.pack(fill="x")
    # Add a button to send the email
    send_button = tk.Button(button_frame, text="Send Email -> Next Preview", command=lambda: [collect_data(), send_email(custom_data, preview_window, customer)], width=140, bg='lightgreen')
    send_button.pack(side="left")
    next_preview_button = tk.Button(button_frame, text="Next Preview", command=lambda: preview_window.destroy(), width=60, bg='yellow')
    next_preview_button.pack(side="right")
    close_button = tk.Button(button_frame, text="Close Program", command=lambda: root.destroy(), width = 60, bg='orange')
    close_button.pack(side="right")

    
    
# Set up the tkinter window
root = tk.Tk()
root.title("WO Processing Tool")

# Create a button to browse files
browse_button = tk.Button(root, text="Select Daily WO File", command=browse_files, height=5, width=50)
browse_button.pack()

# Run the tkinter main loop
root.mainloop()