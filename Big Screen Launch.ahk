#Requires AutoHotkey v2.0
^z::ExitApp





;Change this to Ok or cancel and apphend the entire code block in an result if statement for it
Result := MsgBox("Turn On the TV's and Wait for the image to populate on all 8 screens, then press Ok and relenquish the Mouse", "Launch Monitoring Sites for Big Screen", "1")
if Result = "Ok"

CoordMode "Mouse", "Window"
;Cloud Map
Run "https://weather-radar-live.com/cloud-cover-map/"

;Move to Top Left
if WinWait("Cloud cover map LIVE: ✔️ Where is it cloudy? ⛅️ - Google Chrome", , 10)
{   
    CoordMode "Mouse", "Window"
    Sleep 1000
    WinActivate "Cloud cover map LIVE: ✔️ Where is it cloudy? ⛅️ - Google Chrome"
    WinRestore
    WinGetPos  , , &Width, , "Cloud cover map LIVE: ✔️ Where is it cloudy? ⛅️ - Google Chrome"
    ;Click 1143, 20
    ;Sleep 1000
    ;MouseMove 1913, 203
    
    if (Width = 1936) {
        MsgBox "Drag down to where?"
        Click 1800, 250 ;Ensuring Window is active. (Maybe try Pausing this?)
        Sleep 200
        MouseClickDrag "Left", "1907", "203", "1907", "478", "100"
        Sleep 4000
        MouseMove 1555, 716
        }
    if (Width = 974) {
        MouseClickDrag "Left", "955", "203", "955", "468", "100"
        Sleep 4000
        MouseMove 887, 646
        }
    Sleep 750
    Send "{WheelUp}"
    Sleep 750
    WinMove "-3854", "-2177", 974, 1087
    }
else
{

MsgBox "WinWait timed out Clouds"

}

Sleep 100
SendInput "{Control down}"
SendInput "n"
Sleep 1
SendInput "{Control Up}"
Sleep 1000

;Atlantic Hurricanes
Send "https://www.nhc.noaa.gov/"
Send "{Enter}"

;Move to Bottom Left
if WinWait("National Hurricane Center - Google Chrome", , 10)
{    
WinMove "-2894", "-2177", 974, 1087
Sleep 250
}    
else {

    MsgBox "WinWait timed out NHC"

}


Sleep 1000
SendInput "{Control down}"
SendInput "n"
Sleep 1
SendInput "{Control Up}"
Sleep 1000

;Charter
Send "https://ennexos.sunnyportal.com/login?next=%2F8521253%2Fdashboard"
Send "{Enter}"
if WinWait("Sunny Portal powered by ennexOS - Google Chrome", , 10)
{
CoordMode "Mouse", "Window"
Sleep 1000
WinRestore
Sleep 250
WinMove "-1925", "-1089", 1936, 1096
Sleep 2000
WinMaximize
Sleep 500
Click 208, 488
Sleep 2000
Click 951, 863
;Sleep 1000
;WinMove "-1925", "-1089", 1936, 1096, "Sunny Portal powered by ennexOS - Google Chrome"
}
else {

    MsgBox "Charter Failed"

}

Sleep 1000
SendInput "{Control down}"
SendInput "n"
Sleep 1
SendInput "{Control Up}"
Sleep 1000
WinMove(349, 113, 1401, 802,"New Tab - Google Chrome")
Sleep 1000
Click 1487, 16
Sleep 1000
;Email
Click 1737, 160
;MsgBox "Clicked Missed? grab Coords"






;Also Energy (Needs to be more specific so it opens and logs into Powertrack)
;Send "https://home.alsoenergy.com/login/"
;Send "{Enter}"


CoordMode "Mouse", "Window"

Sleep 1000
;Lily SCADA
Run "G:\Shared drives\O&M\NCC Automations\REMOTE NETWORK SCADA.lnk"
Sleep 5000
;SendInput "AmandaSmith"
;Sleep 0100
;SendInput "{Tab}"
;Sleep 0100
;SendInput "72938"
;Sleep 0100
;SendInput "{Enter}"

;Move to Top Middle Right
if WinWait("US LIL ZONE - PowerStudio Scada", , 15){
    WinActivate "US LIL ZONE - PowerStudio Scada"
    Click 1881, 13
    Sleep 500
    WinMove "-956", "-1710", "640", "480", "US LIL ZONE - PowerStudio Scada"
    Sleep 500
    Click 596, 17
}
else {

    MsgBox "WinWait Timed out!, Lily SCADA"
}



;Run VSS to then open the Multiprocess data pull and hit play
Run "C:\Users\omops\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code.lnk"
Sleep 2000





Run "G:\Shared drives\O&M\NCC Automations\Notification System\#AE API GUI SQL.pyw"
if WinWait("Timestamps", , 10) {
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


    Sleep 1000
if WinExist("NCC Wind Monitor") 
{ 
WinMove("-3845", "-1089", 1936, 1096, "NCC Wind Monitor")
}
else { WinMove("-3845", "-1089", 1936, 1096, "NCC Wind Monitor (Not Responding)")
}

WinMove(36, -594, 500, 584, "Notified events")



Sleep 1000
if WinExist("NCC Wind Monitor") 
{ 
WinActivate("NCC Wind Monitor")
}
else { WinActivate("NCC Wind Monitor (Not Responding)")
}


CoordMode "Mouse", "Screen"
MouseMove 840, 960
WinActivate("Alert Windows Info")
return
}
;If Cancel
else

return



