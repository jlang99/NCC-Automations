#Requires AutoHotkey v2.0
^z::ExitApp

Result := MsgBox("Click ok to launch all monitoring software", "Launch Monitoring Sites for Big Screen", "1")
if Result = "Ok"{

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

Run "G:\Shared drives\O&M\NCC Automations\Notification System\#AE API GUI SQL.pyw"
if WinWait("Timestamps", , 10) {
    WinMove(4767, 46, , , "Lily Update Tool")
    WinMove(3949, 46, , , "Timestamps")
    WinMove(4241, 46, , , "Alert Windows Info")
    WinMove(4200, 330, , , "Personnel On-Site")
    WinMove(2587, 37, , , "Site Data")
    WinMove(5398, 46, , , "Soltage")
    WinMove(5079, 46, , , "NCEMC")
    WinMove(6365, "-305", , , "NARENCO")
    WinMove(7030, "-380", , , "Sol River's Portfolio")
    WinMove(8403, "-365", , , "Harrison Street")
    
   Sleep 1000


Sleep 1000
if WinExist("NCC Weather App") 
{ 
WinActivate("NCC Weather App")
WinMaximize()
}
else { WinActivate("NCC Weather App")
WinMaximize()
}


;Lily Move
WinMove(2398, 746, , , "Notified events")

CoordMode "Mouse", "Screen"
MouseMove 950, 550
WinActivate("Alert Windows Info")
return
}
}
;If Cancel
else {
return
}



