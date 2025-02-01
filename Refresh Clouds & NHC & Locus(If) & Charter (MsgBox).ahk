#Requires AutoHotkey v2.0
AHKMenu := Menu()
AHKMenu.Add("&Open AE API GUI", AEAPIGUI)
AHKMenu.Add("&Close GUI", Close)
AHKMenu.Add("&Postion Time && AE API", Move)
AHKMenu.Add("&Lily SCADA", OpenLily)
AHKMenu.Add("Open &DB", OpenDB)
AHKMenu.Add("Data Pull Script", DataPull)
AHKMenu.Add("Email Notification (Breaker Only)", BreakerNoti)
AHKMenu.Add("Email Notification (Not Working/Don't Use)", EmailNoti)
AHKMenu.Add("Daily Email App", DailyEmail)
AHKMenu.Add("Wind Monitoring App", WindApp)


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
OpenDB(Item, *)
{
   Run "C:\Users\OMOPS\OneDrive - Narenco\Documents\AE API DB.accdb"
}
DataPull(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Notification System\API Data Pull, Multi SQL.py"
}
EmailNoti(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Notification System\Email Notification (Breaker & Inverters).py"
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
   WinClose "SOL River"
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
   WinMove(1095, 204, 308, 314, "Timestamps")
   WinMove(1389, 204, 250, 314, "Alert Windows Info")
   WinMove(1093, 526, , , "Personnel On-Site")
   WinMove(3039, "-874", , , "Soltage")
   WinMove(3421, "-849", , , "NCEMC")
   WinMove(24, "-2132", , , "NARENCO")
   WinMove(1930, "-2133", , , "SOL River")
   WinMove(697, "-2128", , , "Harrison Street")
   WinMove(1302, "-1080", , , "Site Data")
   WinMove(1625, 204, 284, 327, "NCC Desk Functions")

WinMove(36, -594, 500, 584, "Notified events")


}

^!z::AHKMenu.Show

;FUNCITON KEY REASSIGNMENT
f1::{
CoordMode "Mouse", "Screen"
MouseMove 950, 550
}


f2::{
;Cloud Map REFRESH
   CoordMode "Mouse", "Window"
   WinActivate "Cloud cover map LIVE: ✔️ Where is it cloudy? ⛅️ - Google Chrome"
   Sleep 1000
   ;Refresh
   Send "{f5}"

;Zoom on Service area
   Sleep 5000
   Click 879, 874
   Sleep 500
   Send "{WheelUp}"
   ;Sleep 4000
   ;This Won't hold onto the little Map in the Website
   ;MouseClickDrag "L", 871, 843, 623, 812, 100

;Return Mouse to main screen
   Sleep 100
   CoordMode "Mouse", "Screen"
   MouseMove "812", "1051"
Return
}

f3::{
CoordMode "Mouse", "Window"
WinActivate "Cloud cover map LIVE: ✔️ Where is it cloudy? ⛅️ - Google Chrome"
MouseMove 879, 874
}