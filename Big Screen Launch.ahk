#Requires AutoHotkey v2.0
^z::ExitApp

Result := MsgBox("Turn On the TV's and Wait for the image to populate on all 8 screens, then press Ok and relenquish the Mouse", "Launch Monitoring Sites for Big Screen", "1")
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
    WinMove(3297, 785, 284, 327, "NCC Desk Functions")
    WinMove(3624, 477, 429, 314, "Timestamps")
    WinMove(3625, 785, 236, 301, "Alert Windows Info")
    WinMove(2156, 729, , , "Personnel On-Site")
   
    WinMove(2371, "-370", , , "Site Data")
    WinMove(6668, 195, , , "Soltage")
    WinMove(5385, 728, , , "NCEMC")
    WinMove(3000, 24, , , "NARENCO")

    WinMove(4162, "-372", , , "SOL River")

    WinMove(5172, 8, , , "SOL River Continued")

    WinMove(6259, 26, , , "Harrison Street")

   Sleep 1000


Sleep 1000
if WinExist("NCC Wind Monitor") 
{ 
WinActivate("NCC Wind Monitor")
WinMaximize()
}
else { WinActivate("NCC Wind Monitor (Not Responding)")
}

;Lily Move
WinMove(4350, 767, , , "Notified Events")

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



