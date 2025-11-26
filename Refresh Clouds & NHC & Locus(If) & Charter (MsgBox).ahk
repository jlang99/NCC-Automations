#Requires AutoHotkey v2.0
AHKMenu := Menu()
AHKMenu.Add("&Open AE API GUI", AEAPIGUI)
AHKMenu.Add("&Close GUI", Close)
AHKMenu.Add("&Postion Time && AE API", Move)
AHKMenu.Add("&Lily SCADA", OpenLily)
AHKMenu.Add("Data Pull Script", DataPull)
AHKMenu.Add("Email Notification (Breaker Only)", BreakerNoti)
AHKMenu.Add("Daily Email App", DailyEmail)
AHKMenu.Add("NCC Weather App", WindApp)
AHKMenu.Add("WO Input/Customer Notification Tool", WOTool)






;WOGUI := Gui()
;WOGUI.Add("Text", "xm ym", "Work In Progress (Fucntion Not Available)")
;WOGUI.Add("Text", "xm y+5", "Select the Type of WO/Outage")
;WOGUI.AddDropDownList("Choose1 vErrorType", ["Site Trip", "Utility Trip", "Inverter Offline", "Site Stow", "Inverter Performance", "Inverter Comms"])
;WOGUI.Add("Text", "xm y+5", "Select Assign To")
;WOGUI.AddDropDownList("Choose1 Sort vAssigned", ["Brandon","Newman", "Joseph", "Jacob", "Thorne", "Jon", "Isaac", "Parker", "Tom", "Zack"])
;WOGUI.Add("Text", "xm y+5", "Select Site")
;WOGUI.AddDropDownList("Choose1 Sort vLocation", ['Bluebird', 'Cardinal', 'Cherry Blossom', 'Cougar', 'Harrison', 'Hayes', 'Hickory', 'Violet', 'Hickson',
;                    'Jefferson', 'Marshall', 'Ogburn', 'Tedder', 'Thunderhead', 'Van Buren', 'Bulloch 1A', 'Bulloch 1B', 'Elk', 'Duplin',
;                    'Harding', 'Mclean', 'Richmond Cadle', 'Shorthorn', 'Sunflower', 'Upson', 'Warbler', 'Washington', 'Whitehall', 'Whitetail',
;                    'Conetoe', 'Wayne I', 'Wayne II', 'Wayne III', 'Freight Line', 'Holly Swamp', 'PG', 'Bishopville II', 'Gray Fox', 'Wellons'])
;WOGUI.Add("Text", "xm y+5", "Select User")
;WOGUI.AddDropDownList("Choose2 Sort vUser", ["NARENCO", "Joseph", "Jacob"])
;
;SubmitBtn := WOGUI.AddButton("Default w80", "CreateWO")
;SubmitBtn.OnEvent("Click", CreateWO)






WOTool(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Daily Automations\WO Logbook Tool.py"
}

WindApp(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Daily Automations\Wind Weather App.pyw"
}

LilyDown(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Emails\Move Lily Down.ahk"
}

LilyBack(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Emails\Move Lily Back.ahk"
}

OpenLily(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\REMOTE NETWORK SCADA.lnk"
}
AEAPIGUI(Item, *)
{
Run "G:\Shared drives\O&M\NCC Automations\Notification System\#AE API GUI SQL.pyw"
}

DataPull(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Notification System\API Data Pull, Multi SQL.py"
}

BreakerNoti(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Notification System\Email Notification (Breaker).py"
}
DailyEmail(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Emails\NCC Desk Daily Emails Automation.pyw"
}
Close(Item, *)
{
   WinClose "Alert Windows Info"
   WinClose "Sol River's Portfolio"
   WinClose "Harrison Street"
   WinClose "NARENCO"
   WinClose "NCEMC"
   WinClose "Soltage"
   WinClose "Timestamps"
   WinClose "Personnel On-Site"
   WinClose "Site Data"
}

Move(Item, *)
{
    WinMove(3790, 300, , , "NCC Desk Functions")
    WinMove(3074, 300, , , "Timestamps")
    WinMove(3363, 300, , , "Alert Windows Info")
    WinMove(3119, 618, , , "Personnel On-Site")
    WinMove(2198, "-429", , , "Site Data")
    WinMove(2999, "-429", , , "Soltage")
    WinMove(3556, "-429", , , "NCEMC")
    WinMove(4117, "-371", , , "NARENCO")
    WinMove(5172, 8, , , "Sol River's Portfolio")
    WinMove(6259, 26, , , "Harrison Street")
    
WinMove(2398, 746, , , "Notified events")

}


;WO Creation Tool
CreateWO(Item, *)
{
WinMinimize("Timestamps")
WinMinimize("Alert Window")
WinMinimize("NCC Desk Functions")
WinMinimize("Personnel On-Site")
CoordMode "Mouse", "Window"
Click 1694, 184
Sleep 3000
Click 1580, 250
Sleep 1000
Send "3"
Send "{Enter}"
Sleep 1000
Click 1580, 280
Send "Corrective"
Send "{Enter}"
Click 1580, 310
Sleep 1000
Send "Unplanned"
Send "{Enter}"
Click 1580, 340
Sleep 1000

;if vErrorType = "Inverter Offline" {
;   issue := "Equipment Outage"
;}
;elif vErrorType = "Site Trip" {
;   issue := "Site Trip"
;}
;elif vErrorType = "Utility Trip" {
;   issue := "Utility Trip"
;}
;elif vErrorType = "Site Outage" {
;   issue := "Site Trip"
;}

;Send issue   ;Needs If Statement to Filter Type From GUI
Send "{Enter}"
Click 1580, 430
Sleep 1000
Send "Plan"                ;Needs If Statement to Filter Type From GUI
Send "{Enter}"
Click 1580, 610
Sleep 1000
Send "Full"                ;Needs If Statement to Filter Type From GUI
Send "{Enter}"
Click 1580, 640
Sleep 1000
Send "Technician"                ;Needs If Statement to Filter Type From GUI
Send "{Enter}"
Click 1899, 680
Sleep 1000
Send "Thorne"                ;Needs If Statement to Filter Type From GUI
Send "{Enter}"

Click 667, 715
Sleep 1000
Send "Site, Brief Description"                ;Needs If Statement to Filter Type From GUI
Send "{Enter}"
Send "{Tab}"
;Send WO Layout
Send(
    "{Ctrl down}b{Ctrl up}NCC--{Ctrl down}b{Ctrl up}`n"
    "Summary of issue: " A_MM "/" A_DD " - JL - `n"
    "Can we resolve the issue Remotely? (Y or N) N`n"
    "(If applicable) How was the issue resolved?`n"
    "NCC Time: 15 Mins`n"
    "Start Date: " A_MM "/" A_DD "/" A_YYYY "`n"
    "Start Time: " A_Hour ":" A_Min "`n"
    "`n"
    "{Ctrl down}b{Ctrl up}Tech--{Ctrl down}b{Ctrl up}`n"
    "Before photos (Y or N):`n"
    "Reason for equipment fault:`n"
    "During repair photos (Y or N):`n"
    "Summary of work performed:`n"
    "After photos (Y or N):`n"
    "Site Vegetation (Good or Bad)? Be sure to take photos:`n"
    "Time on site:`n"
    "Round trip travel time:`n"
    "Mileage:`n"
    "`n"
    "{Ctrl down}b{Ctrl up}If work order is marked Complete --{Ctrl down}b{Ctrl up}`n"
    "End Date:`n"
    "End Time:`n"
    "`n"
    "{Ctrl down}b{Ctrl up}If Return Trip is needed—{Ctrl down}b{Ctrl up}`n"
    "Tools needed to complete work:`n"
    "Parts to be ordered (Include pics, SN/manufacturer part number):`n"
    "Equipment needed to complete work:`n"
    "Estimated time for repair:"
)

MsgBox("Check for Proper Initials at Summary of Issue:'nInput Proper Start Date and Time and End Time if Applicable")
}



;WOGUI.Show

;HOTKEYS
^!z::AHKMenu.Show

^#r::{
Run "G:\Shared drives\O&M\NCC Automations\Daily Automations\restart.pyw"
}



;FUNCITON KEY REASSIGNMENT
f1::{
CoordMode "Mouse", "Screen"
MouseMove 950, 550
}


f2::{
CoordMode "Mouse", "Screen"
MouseMove 3645, 560
}

f3::{
   CoordMode "Mouse", "Screen"
   MouseMove 165, 1050
   Click 165, 1050
}
F8::AHKMenu.Show
