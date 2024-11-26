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
    WinMove "-3856", "-2143", 1920, 1080
    Sleep 1000
    Send "{f11}"
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
WinMove "-3856", "-1063", "1920", "1080"
Sleep 250
    Send "{f11}"
}    
else {

    MsgBox "WinWait timed out NHC"

}

;Sleep 1000
;SendInput "{Control down}"
;SendInput "n"
;Sleep 1
;SendInput "{Control Up}"
;Sleep 1000

;Issue Tracker
;Sleep 100
;Click 95, 105


;if WinWait("Logbook Open Issues - Google Sheets - Google Chrome", , 10)
;{
;WinMove "-1920", "-1088", "1920", "1080"
;Sleep 250
;    Send "{f11}"
;Sleep 1500
;Click 250, 1058
;}
;else
;Sleep 1000




Sleep 1000
SendInput "{Control down}"
SendInput "n"
Sleep 1
SendInput "{Control Up}"
Sleep 1000

;Google Earth Browser
Send "https://earth.google.com/web/@34.61953748,-81.29366889,429.77805533a,1430789.0199092d,35y,0h,0t,0r/data=OgMKATA"
Send "{Enter}"
;Move to Row 1 Column C
if WinWait("Google Earth - Google Chrome", , 10)
{
WinMove "-1944", "-2184", "1920", "1080"    
    ; Used Google Earth's open on Startup Function the ⭐ This should be a better alternative.
    ;Sleep 1000
    ;Click "279", "427"
    ;Sleep 500
    ;Click "278", "360"
    Sleep 250
    Send "{f11}"
}
else {

    MsgBox "WinWait timed out Google Earth"

}

Sleep 1000
SendInput "{Control down}"
SendInput "n"
Sleep 1000
SendInput "{Control Up}"
Sleep 1000

;CDIA & Cougar LocusNOC
Send "https://locusnoc.datareadings.com/login"
Send "{Enter}"
if WinWait("LocusNOC - Google Chrome", , 10)
{
WinRestore
Sleep 250
WinMove "937", "-1080", "974", "1087"
Sleep 1000
Click 477, 644
}
else {

   MsgBox "LocusNOC login Failed"

}

if WinWait("LocusNOC - Overview - Google Chrome", , 10) {
Sleep 1000
WinRestore
Sleep 250
WinMove "937", "-1080", "974", "1087"
Click 479, 642
}
else {
MsgBox "Locus Noc Move Failed, TRYING AGAIN"
Sleep 300
Click "844", "677"
WinMove "937", "-1080", "974", "1087"
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
WinMove "-23", "-1080", "974", "1087"
Sleep 2000
Click 324, 699
Sleep 200
Send('{Enter}')
;Sleep 1000
;WinMove "-23", "-1080", "974", "1087", "Sunny Portal powered by ennexOS - Google Chrome"
}
else {

    MsgBox "Charter Failed"

}

Sleep 1000
SendInput "{Control down}"
SendInput "n"
Sleep 1
SendInput "{Control Up}"
Sleep 250
WinMove("-8", "-8", 1920, 1048,"New Tab - Google Chrome")
Sleep 1000
WinMaximize("New Tab - Google Chrome")
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
    WinMove "556", "-1832", "640", "480", "US LIL ZONE - PowerStudio Scada"
    Sleep 500
    Click 596, 17
}
else {

    MsgBox "WinWait Timed out!, Lily SCADA"
}



;Run VSS to then open the Multiprocess data pull and hit play
Run "C:\Users\omops\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code.lnk"
Sleep 2000





Run "G:\Shared drives\O&M\NCC Automations\Notification System\#AE API GUI.pyw"
if WinWait("Timestamps", , 10) {
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
    Sleep 1000
if WinExist("NCC Wind Monitor") 
{ 
WinMove("-1944", "-1096", 1936, 1057, "NCC Wind Monitor")
}
else { WinMove("-1944", "-1096", 1936, 1057, "NCC Wind Monitor (Not Responding)")
}

WinMove(757, 498, 500, 600, "Notified events")



Sleep 1000
if WinExist("NCC Wind Monitor") 
{ 
WinActivate("NCC Wind Monitor")
}
else { WinActivate("NCC Wind Monitor (Not Responding)")
}
Click 1858, 16
Sleep 500
WinActivate("LocusNOC - Overview - Google Chrome")
Sleep 500
Click 366, 700
Click 943, 224
Sleep 500






CoordMode "Mouse", "Screen"
MouseMove 840, 960
WinActivate("Alert Windows Info")
return
}
;If Cancel
else

return



