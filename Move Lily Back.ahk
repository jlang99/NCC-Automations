#Requires AutoHotkey v2.0
If WinExist("US LIL ZONE - PowerStudio Scada") {
    WinActivate("US LIL ZONE - PowerStudio Scada")
    Click 1881, 13
    Sleep 500
    WinMove "556", "-1832", "640", "480", "US LIL ZONE - PowerStudio Scada"
    Sleep 500
    Click 596, 17
}