#Requires AutoHotkey v2.0
^z::ExitApp

Result := MsgBox("Wait for atleast 1 Data pull to be successful before clicking ok. Should be the only terminal running and should be communicating `n Pulled Data: Site     | Time Taken: Seconds", "Launch Monitoring Sites for Big Screen", "1")
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
    WinMove(3790, 300, , , "NCC Desk Functions")
    WinMove(3074, 300, , , "Timestamps")
    WinMove(3363, 300, , , "Alert Windows Info")
    WinMove(3119, 618, , , "Personnel On-Site")
    WinMove(2198, "-429", , , "Site Data")
    WinMove(2999, "-422", , , "Soltage")
    WinMove(3556, "-429", , , "NCEMC")
    WinMove(4117, "-371", , , "NARENCO")
    WinMove(5172, 8, , , "Sol River's Portfolio")
    WinMove(6259, 26, , , "Harrison Street")

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



