#Requires AutoHotkey v2.0
AHKMenu := Menu()
AHKMenu.Add("&Open AE API GUI", AEAPIGUI)
AHKMenu.Add("&Close GUI", Close)
AHKMenu.Add("&Postion Time && AE API", Move)
AHKMenu.Add("&Lily SCADA", OpenLily)
AHKMenu.Add("Open &DB", OpenDB)
AHKMenu.Add("Data Pull Script", DataPull)
AHKMenu.Add("Email Notification (set to BA)", EmailNoti)
AHKMenu.Add("Daily Email App", DailyEmail)
AHKMenu.Add("Wind Monitoring App", WindApp)
AHKMenu.Add("AE API Data DB", AEDB)




AEDB(Item, *)
{
   Run "C:\Users\OMOPS\OneDrive - Narenco\Documents\AE API DB.accdb"
   MsgBox("Don't Change the layout of the tables inside this DB")
}

WindApp(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Daily Automations\Wind Weather App (Web Scraping).pyw"
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
Run "G:\Shared drives\O&M\NCC Automations\Notification System\#AE API GUI.pyw"
}
OpenDB(Item, *)
{
   Run "C:\Users\OMOPS\OneDrive - Narenco\Documents\AE API DB.accdb"
}
DataPull(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Notification System\API Data Pull, Multi.py"
}
EmailNoti(Item, *)
{
   Run "G:\Shared drives\O&M\NCC Automations\Notification System\Email Notification (Breaker & Inverters).py"
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
   WinMove(1389, 204, 250, 276, "Alert Windows Info")
   WinMove(1252, 526, , , "Personnel On-Site")
   WinMove(3506, "-621", 335, 585, "Soltage")
   WinMove(3251, "-510", 267, 510, "NCEMC")
   WinMove(3370, "-2153", 464, 1535, "NARENCO")
   WinMove(1916, "-2157", 838, 2060, "SOL River")
   WinMove(2758, "-2154", 540, 1660, "Harrison Street")
   WinMove(1986, "-1030", 364, 1032, "Site Data")
   WinMove(1625, 204, 284, 327, "NCC Desk Functions")

   WinMove(757, 498, 500, 600, "Notified events")
}

^!z::AHKMenu.Show

f2::Send("^x")
f3::Send("^v")
f1::{
CoordMode "Mouse", "Screen"
MouseMove 950, 550
}



f4::{
   Send("^x")
Sleep 100
Send("{Down}")
Send("{Down}")
Sleep 1
Send("^v")
}
f6::{
   IB := InputBox("Input Site Name for DB", "Adding INV Table to DB")
   Site := IB.Value
   lpctIB := InputBox("How many Inverters are there?", "Adding INV Table to DB")
   Loopcnt := lpctIB.Value
   INVct := 1
Loop Loopcnt {
   WinActivate("Access - AE API DB : Database- C:\Users\OMOPS\OneDrive - Narenco\Documents\AE API DB.accdb (Access 2007 - 2016 file format)")
   CoordMode "Mouse", "Window"
   Click 70, 548
   Sleep 300
   Send "^v"
   Sleep 700
   TableName := Site " INV " INVct " Data"
   SendInput(TableName)
   INVct:= ++INVct
   Sleep 500
   WinActivate("Paste Table As")
   Click 61, 104
   Sleep 400
   Send "{Enter}"
   Sleep 700
}



}

f5::{
Result := MsgBox("Does Charter need to be signed in?", "Refresh Charter", 4)
if Result = "Yes"
   {   
   BlockInput "On"
      WinExist("Sunny Portal powered by ennexOS - Google Chrome")
      WinActivate "Sunny Portal powered by ennexOS - Google Chrome"
      CoordMode "Mouse", "Window"
      Sleep 250
      Click "311", "700"
      Sleep 100
      Send('{Enter}')
      Sleep 4000
      Click "311", "700"
      Sleep 300
      Send('{Enter}')
      Click 323, 899

   ;When Scheduled outages occur the Login button moves

   }
else {
   BlockInput "On"
   CoordMode "Mouse", "Window"
}


;Cloud Map REFRESH
   WinActivate "Cloud cover map LIVE: ✔️ Where is it cloudy? ⛅️ - Google Chrome"
   Sleep 250
   Send "{f11}"
   Sleep 1000
   ;Refresh
   Click 99, 63


;Zoom on Service area
   if WinWait("Cloud cover map LIVE: ✔️ Where is it cloudy? ⛅️ - Google Chrome", , 10)
   {
      WinActivate "Cloud cover map LIVE: ✔️ Where is it cloudy? ⛅️ - Google Chrome"
      Sleep 8000
      Click 1552, 769
      Sleep 500
      Send "{WheelUp}"
   }
   else
      MsgBox "error, good luck, lol"

   Sleep 2000
   Send "{f11}"
   Sleep 500


WinActivate "Personnel On-Site"

;Return Mouse to main screen
   Sleep 100
   CoordMode "Mouse", "Screen"
   MouseMove "812", "1051"
   BlockInput "Off"

Return
}

