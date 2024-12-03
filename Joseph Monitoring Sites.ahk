#Requires AutoHotkey v2.0




LilyCBMenu := Menu()
LilyCBMenu.Add("INV 1 Table", CB1)
LilyCBMenu.Add("INV 2 Table", CB2)
LilyCBMenu.Add("INV 3 Table", CB3)
LilyCBMenu.Add("INV 4 Table", CB4)
LilyCBMenu.Add("INV 5 Table", CB5)
LilyCBMenu.Add("INV 6 Table", CB6)
LilyCBMenu.Add("INV 7 Table", CB7)
LilyCBMenu.Add("INV 8 Table", CB8)
LilyCBMenu.Add("INV 9 Table", CB9)
LilyCBMenu.Add("INV 10 Table", CB10)
LilyCBMenu.Add("INV 11 Table", CB11)
LilyCBMenu.Add("INV 12 Table", CB12)
LilyCBMenu.Add("INV 13 Table", CB13)
LilyCBMenu.Add("INV 14 Table", CB14)
LilyCBMenu.Add("INV 15 Table", CB15)
LilyCBMenu.Add("INV 16 Table", CB16)
LilyCBMenu.Add("INV 17 Table", CB17)
LilyCBMenu.Add("INV 18 Table", CB18)
LilyCBMenu.Add("INV 19 Table", CB19)
LilyCBMenu.Add("INV 20 Table", CB20)
LilyCBMenu.Add("INV 21 Table", CB21)
LilyCBMenu.Add("INV 22 Table", CB22)
LilyCBMenu.Add("INV 23 Table", CB23)
LilyCBMenu.Add("INV 24 Table", CB24)
LilyCBMenu.Add("INV 25 Table", CB25)
LilyCBMenu.Add("INV 26 Table", CB26)
LilyCBMenu.Add("INV 27 Table", CB27)
LilyCBMenu.Add("INV 28 Table", CB28)



AHKMenu := Menu()
AHKMenu.Add("OMOPS Websites && Lily Scada", OMOPSsites)
AHKMenu.Add("Joseph's Sites", JosephSites)
AHKMenu.Add("Lily On-Site Personnel Update", LilyPersonnelUpdate)
AHKMenu.Add("Lily Events", LilyExcelTable)
AHKMenu.Add("Lily Updates Notes", LilyNotes)
AHKMenu.Add("Forecasting Report", Forecast)
AHKMenu.Add("Open Logbook", Logbook)
AHKMenu.Add("Tracker Check", Tracker)
AHKMenu.Add("CB Check", CBCheck)
AHKMenu.Add("Shift Summary Email App", EmailApp)
AHKMenu.Add("Change Shift Summary Query Start", ShiftSumStart)
AHKMenu.Add("WO Input Doc", WOInput)
AHKMenu.Add("NCC WO's Input && Notify", WOInputTask)



;Setting Icons
AHKMenu.SetIcon("Lily On-Site Personnel Update", "G:\Shared drives\O&M\NCC Automations\Icons\excel-file.ico")
AHKMenu.SetIcon("OMOPS Websites && Lily Scada", "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", 1)
AHKMenu.SetIcon("Joseph's Sites", "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", 1)
AHKMenu.SetIcon("Lily Events", "G:\Shared drives\O&M\NCC Automations\Icons\excel-file.ico")
AHKMenu.SetIcon("Open Logbook", "G:\Shared drives\O&M\NCC Automations\Icons\Logbook.ico")
AHKMenu.SetIcon("Forecasting Report", "G:\Shared drives\O&M\NCC Automations\Icons\Forecast.ico")
AHKMenu.SetIcon("Lily Updates Notes", "G:\Shared drives\O&M\NCC Automations\Icons\Notes.ico")
AHKMenu.SetIcon("Tracker Check", "G:\Shared drives\O&M\NCC Automations\Icons\tracker_3KU_icon.ico")
AHKMenu.SetIcon("CB Check", "G:\Shared drives\O&M\NCC Automations\Icons\tracker_3KU_icon.ico")

;Adobe Help
f1::{
    MouseGetPos &x, &y
    Sleep 100
    CoordMode "Mouse", "Window"
    Click 319, 350
    Sleep 200
    MouseMove x, y
}
!Down::{
    MouseMove 0, 9, , "R"
}
!Up::{
    MouseMove 0, -9, , "R"
}
;Tracker input INV associated Help
;NumpadEnter::Send(", ")





;ED Input
^#NumpadEnter::Send(A_Hour ":" A_Min)
^#NumpadAdd::Send(A_MM "/" A_DD "/" A_YYYY)

;Menu
^<!z::AHKMenu.Show
^+!z::LilyCBMenu.Show

;WO Help
^#j::Send(
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

;Emails
^#b::Send("brandon.arrowood@narenco.com")
^#m::Send("joseph.lang@narenco.com")

;Logbook Help
^#n::Send("NARENCO|")
^#Numpad0::Send("0800")

^#+v::{
    Send("SOUTH TEXAS CURVING")
    Send('{Tab}')
    Send("3")
    Send('{Tab}')
    Send("0800")
    Send('{Tab}')
    Send("Yes")
    CoordMode "Mouse", "Window"
    Sleep 200
    Click 251, 561
    Sleep 250
    Send("VEGETATION MANAGEMENT")
    Send('{Tab}')
    Send('{Tab}')
    Send("Mowing")
}

^#v::{
    WinActivate("Access - NCC 039 : Database- G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb (Access 2007 - 2016 file format)")
    CoordMode "Mouse", "Window"
    Sleep 200
    Click 251, 561
    Sleep 250
    Send("VEGETATION MANAGEMENT")
    Send('{Tab}')
    Send('{Tab}')
    Send("Mowing")
}


^#c::{
WinActivate("Enter Parameter Value")
CoordMode "Mouse", "Window"
Sleep 100
Click 195, 99
Sleep 250
Click
Sleep 250
WinActivate("Access - NCC 039 : Database- G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb (Access 2007 - 2016 file format)")
Click 1113, 531
Sleep 200
Click 1116, 771
Sleep 200
Click 1103, 995
}

^#r::{
Run "G:\Shared drives\O&M\NCC Automations\Daily Automations\restart.pyw"
}



f8::{
    CoordMode "Mouse", "Screen"
    Click 251, 1052
    WinWait("Google Chrome", , 10)
        CoordMode "Mouse", "Window"
        WinActivate("Google Chrome")
        Sleep 1500
        Click 428, 341
        Sleep 1500
        WinMove( 1916, 1072, 1936, 1048, "New Tab - Google Chrome")
        Sleep 500
        Click 74, 111
        Sleep 500
    CoordMode "Mouse", "Screen"
    Click 251, 1052
    WinWait("Google Chrome", , 10)
        Sleep 2000
        Click 689, 470
        Sleep 1500 
        Run "https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox"
        Sleep 100
        Run "https://tsasolaroperations.dominionenergyse.com/#"
        Sleep 100
        Run "https://drive.google.com/drive/my-drive"
        Sleep 100
        Run "https://docs.google.com/spreadsheets/d/1wqtrYMKMUtr6WXP3KULoyLfbSPHLiGP0wOBSKq9OH00/edit#gid=0"
        Sleep 100
        Run "https://keep.google.com/#label/Investigate%20Later"
        Sleep 100
        Run "https://www.suncalc.org/"
        Sleep 100
        Run "https://www.wbtv.com/weather/"
        Sleep 100
        If(A_DDDD = "Sunday" or A_DDDD = "Saturday")
        Run "https://weather-radar-live.com/cloud-cover-map/"
        Sleep 100
        Run "https://app.raptormaps.com/solar-sites"
        Sleep 100
        Run "https://academy.viewpoint.com/learn/signin"
        Sleep 500
        WinMove( 1916, 1072, 1936, 1048)

    Click 251, 1052
    WinWait("Google Chrome", , 10)
        Sleep 5000
        Click 870, 662
        Sleep 1500
        
        ;Chrome Sites
        Run "chrome.exe https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox"
        Sleep 200
        Run "https://app-next.gamechangesolar.com/layouts?view=list"
        Sleep 200
        Run "https://www.suntrack.services/en/"
        Sleep 200
        Run "https://apps.alsoenergy.com/powertrack/S63226/overview/dashboard"
        Sleep 200
        Run "https://elks.aderisenergy.com/dashboard"
        Sleep 200
        Run "https://ennexos.sunnyportal.com/8521253/dashboard"
        Sleep 200
        Run "https://x46.emaint.com/wc.dll?X3~emproc~x3Hubv2#/WORK/__new/ADD"
        Sleep 200
        Run "https://events.x-elio.com/list.php"
        Sleep 200
        Run "https://earth.google.com/web/@34.82626483,-79.87550174,-39.97826276a,1240758.82427394d,30.00018764y,0h,0t,0r/data=CgRCAggBMikKJwolCiExQ19lVkZBUnRWMi00NlJzbVZkSTM2dHVVay1CVnMxVHIgAToDCgEwQgIIAEoICJXPs44HEAE"
        Sleep 200
        Run "https://drive.google.com/drive/u/0/my-drive"
        Sleep 200
        ;Underperformances
        Run "https://docs.google.com/spreadsheets/d/1sp8SJNbMf0AhtEn2-WZmCUKHCGZrqRDgNFNOtoUkuHE/edit?gid=1107988056#gid=1107988056"
        Sleep 200
        ;Tracker Issue sheet
        Run "https://docs.google.com/spreadsheets/d/1EzgmTCAAkCuOIVVRTUtNOp_jXdZpT7Hob-DUhEuYmrE/edit?gid=1107988056#gid=1107988056"
        Sleep 200
        ; CB Issue Sheet
        Run "https://docs.google.com/spreadsheets/d/1RGUwARwDdfDoC8VcNQgb5KC6gPOA-rcMaPsu6vlzg9o/edit?gid=1611132963#gid=1611132963"
        Sleep 200
        ;Site List
        Run "https://docs.google.com/spreadsheets/d/1GdO46Jt304OLf-H-dGJ3MQQUbGbd8lEUAuOewhDdIWM/edit#gid=1126541456"
        Sleep 200
        ; NARENCO Connections sheet
        Run "https://docs.google.com/spreadsheets/d/1YTd0c_VLAGog1NTxo0VGzAdGFUhs9oquY1WJFz2NJA0/edit#gid=0"
        Sleep 200
        ; Lily CB Check
        If(A_DDDD = "Sunday" or A_DDDD = "Saturday")
        Run "https://docs.google.com/spreadsheets/d/1I4UfcXZYG87b4JmBahjgO4EADph-z4pCk6R9SS7vuW0/edit#gid=2071680363"
        Sleep 200
        ; Lily On Site List
        If(A_DDDD = "Tuesday" or A_DDDD = "Wednesday" or A_DDDD = "Thursday" or A_DDDD = "Friday" or A_DDDD = "Saturday")
        Run "https://docs.google.com/spreadsheets/d/1Uq8LSya5w6xiAbFhqeCa1upT2Zs-ueQmYFP6cxNkOkI/edit#gid=0"
        Sleep 200
        Run "https://keep.google.com/#label/1.%20Handover"
        Sleep 2000
        CoordMode "Mouse", "Window"
        Sleep 500
        Click 1862, 23
        Sleep 3000
        WinMove( 2056, "-6", 1656, 1013, "Google Keep - Google Chrome")
        Sleep 500
        Click 1574, 15

        Sleep 10000
        BlockInput 1
        ;Lily Scada
        #Warn All, Off
        Lily := "C:\AppletScada\AppletScada_x4.jar 12.42.18.10:22000 user:AmandaSmith password:72938 multipleinstance"
        Run Lily
        CoordMode "Mouse", "Window"
        WinWait("US LIL ZONE - PowerStudio Scada", , 15)
        Winactivate("US LIL ZONE - PowerStudio Scada")
        Click 1880, 11
        Sleep 1000
        If(A_DDDD = "Sunday" or A_DDDD = "Saturday")
            WinMove 409, 286, 640, 480
        Else
            WinMove 347, 1323, 640, 480
        Sleep 2000
        Click 594, 16
        WinMove(898, 1357, 500, 600, "Notified events")
        Sleep 250
        WinMinimize("Notified events")

        BlockInput 0

        Sleep 1000

        Run "G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb"
        Sleep 5000
        If(A_DDDD = "Sunday" or A_DDDD = "Saturday") {
            CoordMode "Mouse", "Window"
            WinActivate "Access - NCC 039 : Database- G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb (Access 2007 - 2016 file format)"
            Click 1852, 17
            Sleep 100
            WinMove 149, 1205, 1440, 752, "Access - NCC 039 : Database- G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb (Access 2007 - 2016 file format)"
            Sleep 250
            Click 1359, 17
        }

        Sleep 1000

        Run "G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\ForecastReport 005.xlsm"
        WinWait('ForecastReport 005.xlsm - Excel', , 10)
        Sleep 8000
        WinActivate("ForecastReport 005.xlsm - Excel")
        Sleep 250 
        WinMaximize("ForecastReport 005.xlsm - Excel")
        Sleep 1000
        CoordMode "Mouse", "Window"
        Click 154, 1000
        Sleep 250
        Click 374, 277
        WinWait("RUN FORECAST PROCEDURE", , 10)
        Click 190, 120
}

WOInputTask(Item, *)
{
    Run "G:\Shared drives\O&M\NCC Automations\Daily Automations\WO Logbook Tool.py"
}
WOInput(Item, *)
{
    Run "G:\Shared drives\Narenco Projects\O&M Projects\NCC\Procedures\Tech WO Input.docx"
}

EmailApp(Item, *)
{
    Run "G:\Shared drives\O&M\NCC Automations\Emails\NCC Desk Daily Emails Automation.pyw"
}
ShiftSumStart(Item, *)
{
    Run "G:\Shared drives\O&M\NCC Automations\Emails\Shift Summary Start.txt"
    MsgBox "Input the last entries LogID# from the previous shift"
}

CBCheck(Item, *)
{
    Run "G:\Shared drives\O&M\NCC Automations\Daily Automations\CB Check.py"
}


Tracker(Item, *)
{
    Run "G:\Shared drives\O&M\NCC Automations\Daily Automations\Tracker Check All Sites.py"
}


Logbook(Item, *)
{
    Run "G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb"
}

Forecast(Item, *)
{
    Run "G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\ForecastReport 005.xlsm"
    WinWait('ForecastReport 005.xlsm - Excel', , 10)
    Sleep 4000
    CoordMode "Mouse", "Window"
    Click 154, 1000
    Sleep 250
    Click 374, 277
    WinWait("RUN FORECAST PROCEDURE", , 10)
    Click 190, 120
}


LilyExcelTable(Item, *)
{
    Run "https://docs.google.com/spreadsheets/d/1byNW58NbhuDotmdzJEFplczsXLAP0iMyS9YVU5FqQgk/edit#gid=0"
}
LilyNotes(Item, *)
{
    Run "G:\Shared drives\O&M\NCC Automations\Emails\Lily Updates.txt"
}

OpenIssueTracker(Item, *)
{
Run "C:\Users\OMOPS.AzureAD\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Open Issues Update.py"
}


;Lily CB Menu Generate Tables
CB1(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 158, 151, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 183, 2
Sleep 500
;Selecting Mod1
Click 210, 206, 2

WinWait("Variables selection (US LIL INVERTER 01 STRING MOD1)", , 10)
Sleep 1000
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 01 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 500
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 372

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 158, 151, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 183, 2
Sleep 500
;Add Mod 2
Click 210, 225, 2

WinWait "Variables selection (US LIL INVERTER 01 STRING MOD2)", , 10
Sleep 1000
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;first Row
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 81, 442
WinWait "Graph - US LIL INVERTER 01 STRING MOD1, US LIL INVERTER 01 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB2(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 158, 151, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 183, 2
Sleep 500
;Selecting Mod1
Click 244, 240, 2

WinWait("Variables selection (US LIL INVERTER 02 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 02 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 372

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 158, 151, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 183, 2
Sleep 500
;Add Mod 2
Click 250, 253, 2

WinWait "Variables selection (US LIL INVERTER 02 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 81, 442
WinWait "Graph - US LIL INVERTER 02 STRING MOD1, US LIL INVERTER 02 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB3(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 135, 184, 2
Sleep 500
;Folder 3 (String Mod)
Click 171, 214, 2
Sleep 500
;Selecting Mod1
Click 239, 241, 2

WinWait("Variables selection (US LIL INVERTER 03 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 03 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 135, 184, 2
Sleep 500
;Folder 3 (String Mod)
Click 171, 214, 2
Sleep 500
;Selecting Mod2
Click 259, 258, 2

WinWait "Variables selection (US LIL INVERTER 03 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 03 STRING MOD1, US LIL INVERTER 03 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB4(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 135, 184, 2
Sleep 500
;Folder 3 (String Mod)
Click 171, 214, 2
Sleep 500
;Selecting Mod1
Click 269, 273, 2

WinWait("Variables selection (US LIL INVERTER 04 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 04 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 372

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 135, 184, 2
Sleep 500
;Folder 3 (String Mod)
Click 171, 214, 2
Sleep 500
;Selecting Mod2
Click 252, 284, 2

WinWait "Variables selection (US LIL INVERTER 04 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 89, 443
WinWait "Graph - US LIL INVERTER 04 STRING MOD1, US LIL INVERTER 04 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB5(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 157, 218, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 248, 2
Sleep 500
;Selecting Mod1
Click 238, 274, 2

WinWait("Variables selection (US LIL INVERTER 05 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 05 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 157, 218, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 248, 2
Sleep 500
;Selecting Mod2
Click 238, 286, 2

WinWait "Variables selection (US LIL INVERTER 05 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 05 STRING MOD1, US LIL INVERTER 05 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB6(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 157, 218, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 248, 2
Sleep 500
;Selecting Mod1
Click 238, 302, 2

WinWait("Variables selection (US LIL INVERTER 06 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 06 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 157, 218, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 248, 2
Sleep 500
;Selecting Mod2
Click 238, 317, 2

WinWait "Variables selection (US LIL INVERTER 06 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 06 STRING MOD1, US LIL INVERTER 06 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB7(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 164, 247, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 278, 2
Sleep 500
;Selecting Mod1
Click 238, 304, 2

WinWait("Variables selection (US LIL INVERTER 07 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 07 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 164, 247, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 278, 2
Sleep 500
;Selecting Mod2
Click 238, 317, 2

WinWait "Variables selection (US LIL INVERTER 07 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 07 STRING MOD1, US LIL INVERTER 07 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB8(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 164, 247, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 278, 2
Sleep 500
;Selecting Mod1
Click 238, 338, 2

WinWait("Variables selection (US LIL INVERTER 08 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 08 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 164, 247, 2
Sleep 500
;Folder 3 (String Mod)
Click 144, 278, 2
Sleep 500
;Selecting Mod2
Click 238, 351, 2

WinWait "Variables selection (US LIL INVERTER 08 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 08 STRING MOD1, US LIL INVERTER 08 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB9(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 280, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod1
Click 202, 305, 2

WinWait("Variables selection (US LIL INVERTER 09 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 09 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 373

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 280, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod2
Click 171, 324, 2

WinWait "Variables selection (US LIL INVERTER 09 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 444
WinWait "Graph - US LIL INVERTER 09 STRING MOD1, US LIL INVERTER 09 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB10(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 280, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod1
Click 238, 336, 2

WinWait("Variables selection (US LIL INVERTER 10 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 10 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 280, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod2
Click 171, 350, 2

WinWait "Variables selection (US LIL INVERTER 10 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 10 STRING MOD1, US LIL INVERTER 10 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB11(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 310, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod1
Click 238, 306, 2

WinWait("Variables selection (US LIL INVERTER 11 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 11 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 42, 429

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 310, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod2
Click 171, 319, 2

WinWait "Variables selection (US LIL INVERTER 11 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 85, 500
WinWait "Graph - US LIL INVERTER 11 STRING MOD1, US LIL INVERTER 11 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB12(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 310, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod1
Click 202, 335, 2

WinWait("Variables selection (US LIL INVERTER 12 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 12 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 373

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 310, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod2
Click 171, 354, 2

WinWait "Variables selection (US LIL INVERTER 12 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 444
WinWait "Graph - US LIL INVERTER 12 STRING MOD1, US LIL INVERTER 12 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB13(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 152, 343, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 329, 2
Sleep 500
;Selecting Mod1
Click 238, 351, 2

WinWait("Variables selection (US LIL INVERTER 13 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 13 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 152, 343, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 329, 2
Sleep 500
;Selecting Mod2
Click 171, 366, 2

WinWait "Variables selection (US LIL INVERTER 13 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 13 STRING MOD1, US LIL INVERTER 13 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB14(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 375, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod1
Click 202, 302, 2

WinWait("Variables selection (US LIL INVERTER 14 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 14 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 373

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 375, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod2
Click 171, 324, 2

WinWait "Variables selection (US LIL INVERTER 14 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 444
WinWait "Graph - US LIL INVERTER 14 STRING MOD1, US LIL INVERTER 14 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB15(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 375, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod1
Click 238, 332, 2

WinWait("Variables selection (US LIL INVERTER 15 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 15 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 375, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 280, 2
Sleep 500
;Selecting Mod2
Click 171, 350, 2

WinWait "Variables selection (US LIL INVERTER 15 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 15 STRING MOD1, US LIL INVERTER 15 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB16(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 408, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 312, 2
Sleep 500
;Selecting Mod1
Click 238, 336, 2

WinWait("Variables selection (US LIL INVERTER 16 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 16 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 408, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 312, 2
Sleep 500
;Selecting Mod2
Click 171, 350, 2

WinWait "Variables selection (US LIL INVERTER 16 STRING MOD2)", , 10
Sleep 500
;All
Click 679, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 16 STRING MOD1, US LIL INVERTER 16 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB17(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 408, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 312, 2
Sleep 500
;Selecting Mod1
Click 238, 368, 2

WinWait("Variables selection (US LIL INVERTER 17 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 17 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500
;Folder 2 INV Selection
Click 175, 408, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 312, 2
Sleep 500
;Selecting Mod2
Click 171, 384, 2

WinWait "Variables selection (US LIL INVERTER 17 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 17 STRING MOD1, US LIL INVERTER 17 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB18(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 325, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 299, 2
Sleep 500

;Selecting Mod1
Click 238, 319, 2

WinWait("Variables selection (US LIL INVERTER 18 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Third Blank
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 18 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 343

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 325, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 299, 2
Sleep 500

;Selecting Mod2
Click 171, 334, 2

WinWait "Variables selection (US LIL INVERTER 18 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 18 STRING MOD1, US LIL INVERTER 18 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB19(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 325, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 299, 2
Sleep 500

;Selecting Mod1
Click 238, 354, 2

WinWait("Variables selection (US LIL INVERTER 19 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 19 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 399

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 325, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 299, 2
Sleep 500

;Selecting Mod2
Click 171, 369, 2

WinWait "Variables selection (US LIL INVERTER 19 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 100
;Row 5
Click 416, 340
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 100
;Click OK
Click 90, 475
WinWait "Graph - US LIL INVERTER 19 STRING MOD1, US LIL INVERTER 19 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB20(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 295, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 295, 2
Sleep 500

;Selecting Mod1
Click 238, 320, 2

WinWait("Variables selection (US LIL INVERTER 20 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 20 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 372

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 295, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 295, 2
Sleep 500

;Selecting Mod2
Click 171, 339, 2

WinWait "Variables selection (US LIL INVERTER 20 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 200
;Click OK
Click 102, 435
WinWait "Graph - US LIL INVERTER 20 STRING MOD1, US LIL INVERTER 20 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB21(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 256, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 290, 2
Sleep 500

;Selecting Mod1
Click 238, 317, 2

WinWait("Variables selection (US LIL INVERTER 21 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 21 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 342

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 256, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 290, 2
Sleep 500

;Selecting Mod2
Click 171, 335, 2

WinWait "Variables selection (US LIL INVERTER 21 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 200
;Click OK
Click 102, 415
WinWait "Graph - US LIL INVERTER 21 STRING MOD1, US LIL INVERTER 21 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB22(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 256, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 290, 2
Sleep 500

;Selecting Mod1
Click 238, 350, 2

WinWait("Variables selection (US LIL INVERTER 22 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 22 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 372

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 256, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 290, 2
Sleep 500

;Selecting Mod2
Click 171, 365, 2

WinWait "Variables selection (US LIL INVERTER 22 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 200
;Click OK
Click 102, 444
WinWait "Graph - US LIL INVERTER 22 STRING MOD1, US LIL INVERTER 22 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB23(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 165, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 195, 2
Sleep 500

;Selecting Mod1
Click 238, 224, 2

WinWait("Variables selection (US LIL INVERTER 23 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 23 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 342

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 165, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 195, 2
Sleep 500

;Selecting Mod2
Click 171, 236, 2

WinWait "Variables selection (US LIL INVERTER 23 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 200
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 23 STRING MOD1, US LIL INVERTER 23 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB24(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 165, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 195, 2
Sleep 500

;Selecting Mod1
Click 238, 252, 2

WinWait("Variables selection (US LIL INVERTER 24 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 24 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 342

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 165, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 195, 2
Sleep 500

;Selecting Mod2
Click 171, 269, 2

WinWait "Variables selection (US LIL INVERTER 24 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 200
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 24 STRING MOD1, US LIL INVERTER 24 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB25(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 228, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 258, 2
Sleep 500

;Selecting Mod1
Click 238, 287, 2

WinWait("Variables selection (US LIL INVERTER 25 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 25 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 342

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 228, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 258, 2
Sleep 500

;Selecting Mod2
Click 171, 303, 2

WinWait "Variables selection (US LIL INVERTER 25 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 200
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 25 STRING MOD1, US LIL INVERTER 25 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB26(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 295, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 295, 2
Sleep 500

;Selecting Mod1
Click 238, 355, 2

WinWait("Variables selection (US LIL INVERTER 26 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 26 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 342

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 295, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 295, 2
Sleep 500

;Selecting Mod2
Click 171, 370, 2

WinWait "Variables selection (US LIL INVERTER 26 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 200
;Click OK
Click 102, 415
WinWait "Graph - US LIL INVERTER 26 STRING MOD1, US LIL INVERTER 26 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB27(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 135, 200, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 231, 2
Sleep 500

;Selecting Mod1
Click 238, 254, 2

WinWait("Variables selection (US LIL INVERTER 27 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 27 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 372

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 135, 200, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 231, 2
Sleep 500

;Selecting Mod2
Click 171, 271, 2

WinWait "Variables selection (US LIL INVERTER 27 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 100
;Row 5
Click 416, 340
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 200
;Click OK
Click 102, 444
WinWait "Graph - US LIL INVERTER 27 STRING MOD1, US LIL INVERTER 27 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

CB28(Item, *)
{
BlockInput 1
CoordMode "Mouse", "Window"
;Open Graph
Click 557, 69, 2

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 228, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 258, 2
Sleep 500

;Selecting Mod1
Click 238, 315, 2

WinWait("Variables selection (US LIL INVERTER 28 STRING MOD1)", , 10)
Sleep 500
;All
Click 701, 65
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;First Blank
Click 407, 478
Sleep 100
;Second Blank
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click Ok
Click 237, 572

WinWait "Graph - US LIL INVERTER 28 STRING MOD1 - PowerStudio Scada", , 10
;Properties to Add Second Set of Inputs
Click 835, 70

WinWait "Graph properties", , 10
Sleep 100
;3 Lines to Wheel Down to + Button
Click 312, 175, 0
Sleep 100
Click "WD", 20 
Sleep 500
;Add Button
Click 39, 342

WinWait("Device selection", , 10)
Sleep 500
;Folder 1 
Click 107, 91, 2
Sleep 500

;Scroll to More
Click 358, 168, 0
Sleep 100
Click "WD", 20 
Sleep 500

;Folder 2 INV Selection
Click 175, 228, 2
Sleep 500
;Folder 3 (String Mod)
Click 160, 258, 2
Sleep 500

;Selecting Mod2
Click 171, 333, 2

WinWait "Variables selection (US LIL INVERTER 28 STRING MOD2)", , 10
Sleep 500
;All
Click 693, 66
Sleep 100
;Bottom
Click 380, 512
Sleep 100
;Row 1
Click 407, 478
Sleep 100
;row 2
Click 412, 441
Sleep 100
;Row 3
Click 416, 407
Sleep 100
;Row 4
Click 416, 375
Sleep 1000
;Click OK
Click 237, 572
Sleep 500

WinActivate "Graph properties"
Sleep 200
;Click OK
Click 102, 414
WinWait "Graph - US LIL INVERTER 28 STRING MOD1, US LIL INVERTER 28 STRING MOD2 - PowerStudio Scada", , 10
;Open Table
Send "!b"
BlockInput 0
}

OMOPSsites(Item, *)
{
;Chrome Sites
Run "chrome.exe https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox"
Sleep 200
Run "https://app-next.gamechangesolar.com/layouts?view=list"
Sleep 200
Run "https://www.suntrack.services/en/"
Sleep 200
Run "https://apps.alsoenergy.com/powertrack/S63226/overview/dashboard"
Sleep 200
Run "https://elks.aderisenergy.com/dashboard"
Sleep 200
Run "https://ennexos.sunnyportal.com/8521253/dashboard"
Sleep 200
Run "https://x46.emaint.com/wc.dll?X3~emproc~x3Hubv2#/WORK/__new/ADD"
Sleep 200
Run "https://events.x-elio.com/list.php"
Sleep 200
Run "https://earth.google.com/web/@34.82626483,-79.87550174,-39.97826276a,1240758.82427394d,30.00018764y,0h,0t,0r/data=CgRCAggBMikKJwolCiExQ19lVkZBUnRWMi00NlJzbVZkSTM2dHVVay1CVnMxVHIgAToDCgEwQgIIAEoICJXPs44HEAE"
Sleep 200
Run "https://drive.google.com/drive/u/0/my-drive"
Sleep 200
;Underperformances
Run "https://docs.google.com/spreadsheets/d/1sp8SJNbMf0AhtEn2-WZmCUKHCGZrqRDgNFNOtoUkuHE/edit?gid=1107988056#gid=1107988056"
Sleep 200
;Tracker Issue sheet
Run "https://docs.google.com/spreadsheets/d/1EzgmTCAAkCuOIVVRTUtNOp_jXdZpT7Hob-DUhEuYmrE/edit?gid=1107988056#gid=1107988056"
Sleep 200
; CB Issue Sheet
Run "https://docs.google.com/spreadsheets/d/1RGUwARwDdfDoC8VcNQgb5KC6gPOA-rcMaPsu6vlzg9o/edit?gid=1611132963#gid=1611132963"
Sleep 200
;Site List
Run "https://docs.google.com/spreadsheets/d/1GdO46Jt304OLf-H-dGJ3MQQUbGbd8lEUAuOewhDdIWM/edit#gid=1126541456"
Sleep 200
; NARENCO Connections sheet
Run "https://docs.google.com/spreadsheets/d/1YTd0c_VLAGog1NTxo0VGzAdGFUhs9oquY1WJFz2NJA0/edit#gid=0"
Sleep 200
; Lily CB Check
If(A_DDDD = "Sunday" or A_DDDD = "Saturday")
Run "https://docs.google.com/spreadsheets/d/1I4UfcXZYG87b4JmBahjgO4EADph-z4pCk6R9SS7vuW0/edit#gid=2071680363"
Sleep 200
; Lily On Site List
If(A_DDDD = "Tuesday" or A_DDDD = "Wednesday" or A_DDDD = "Thursday" or A_DDDD = "Friday" or A_DDDD = "Saturday")
Run "https://docs.google.com/spreadsheets/d/1Uq8LSya5w6xiAbFhqeCa1upT2Zs-ueQmYFP6cxNkOkI/edit#gid=0"
Sleep 200
Run "https://keep.google.com/#label/1.%20Handover"



Sleep 10000
BlockInput 1
;Lily Scada
#Warn All, Off
Lily := "C:\AppletScada\AppletScada_x4.jar 12.42.18.10:22000 user:AmandaSmith password:72938 multipleinstance"
Run Lily
CoordMode "Mouse", "Window"
WinWait("US LIL ZONE - PowerStudio Scada", , 15)
Winactivate("US LIL ZONE - PowerStudio Scada")
Click 1880, 11
Sleep 1000
WinMove 347, 1323, 640, 480
Sleep 2000
Click 594, 16
WinMinimize("Notified events")
BlockInput 0
}

JosephSites(Item, *)
{
Run "https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox"
Sleep 100
Run "https://tsasolaroperations.dominionenergyse.com/#"
Sleep 100
Run "https://drive.google.com/drive/my-drive"
Sleep 100
Run "https://docs.google.com/spreadsheets/d/1wqtrYMKMUtr6WXP3KULoyLfbSPHLiGP0wOBSKq9OH00/edit#gid=0"
Sleep 100
Run "https://keep.google.com/#label/Investigate%20Later"
Sleep 100
Run "https://www.suncalc.org/"
Sleep 100
Run "https://www.wbtv.com/weather/"
Sleep 100
If(A_DDDD = "Sunday" or A_DDDD = "Saturday")
Run "https://weather-radar-live.com/cloud-cover-map/"
Sleep 100
Run "https://app.raptormaps.com/solar-sites"
Sleep 100
Run "https://academy.viewpoint.com/learn/signin"
}
LilyPersonnelUpdate(Item, *)
{
    Run "G:\Shared drives\O&M\NCC Automations\Daily Automations\Lily Personnel Update.py"
}


LilyPersonnelUpdateOld(Item, *)
{
BlockInput 1
WinActivate "Access - NCC 039 : Database- G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb (Access 2007 - 2016 file format)"
CoordMode "Mouse", "Window"
Sleep 100
MouseClick "Left", 364, 248
Sleep 4000
Click "56", "200"
Sleep 100
Send "^c"
Sleep 1000
Click "325", "181"
Sleep 500
Run "https://docs.google.com/spreadsheets/d/15g5O5Xd9MobMkb9-COR8JlovFoIN0qIlopQJNzKpNrY/edit#gid=0"
Sleep 8000
Send "^v"
Sleep 100
BlockInput 0
MsgBox "You have updated X-Elio on all Site Access entries at Lily from yesterday Only! Check the Lily On Site Personnel/Contractor Sheet for errors", "Close the Google Sheet"
AHKMenu.ToggleCheck("&Lily On-Site Personnel Update")
JosephMenu.ToggleCheck("&Lily On-Site Personnel Update")
}

Name()
{
BlockInput 1
CoordMode "Mouse", "Screen"
Click "1700", "1050"
WinWait "Cisco AnyConnect Secure Mobility Client", , "10"
CoordMode "Mouse", "Window"
Click "335", "112"
WinWait "Cisco AnyConnect | secnet.alsoenergy.com", , "10"
Send "Nw7gQGyuLcqk0UNw"
Sleep 500
Send "{Enter}"
BlockInput 0
}
^>!z::Name()
ShiftSum(Item, *) {
    Run "C:\Users\OMOPS.AzureAD\OneDrive - Narenco\Documents\AutoHotkey\#Testing Shift Summary Email Send.py"
}

ShiftSums(Item, *)
{
BlockInput 1
SendMode "Input"
WinActivate "Untitled - Paint"
Sleep 100
Send "^c"
Sleep 100
WinActivate "Inbox - omops@narenco.com - Narenco Mail - Google Chrome"
Sleep 100
CoordMode "Mouse", "Window"
Sleep 5000
Click 277, 194
SLeep 100
Click 160, 220
Sleep 5000

Send "Shift"
Sleep 1500
Send "{Enter}"
Sleep 1000
Send "{Tab}"
Sleep 2000
SendInput FormatTime(,"MM/dd/yy")
Sleep 100
Send " NCC Shift Summary"
Sleep 100
Send "{Tab}"
Sleep 100
Send "^v"
Sleep 3000
BlockInput 0

;Result1 :=MsgBox("Is it before 18:00?",, "4")
;If Result1 = "Yes"
;{
;BlockInput 1
;Section to Schedule send at Blank PM
;Click 1360, 1006
;Sleep 500
;Click 1360, 969
;Sleep 500
;Click 940, 680
;Sleep 2000
;Click 1068, 521
;Sleep 100
;Send "^a"
;Sleep 100
;SendInput "1800"
;Click 1129, 741
;BlockInput 0
;MsgBox "Scheduled for 05:45"
;}
;Else
;MsgBox "Hit Send"
return
}


