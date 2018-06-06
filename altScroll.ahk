#SingleInstance, force
AutoReload()
TrayTip, %A_ScriptName%, Loaded
; #NoTrayIcon
GroupAdd, altScrollIgnore, ahk_class Chrome_WidgetWin_1
GroupAdd, altScrollIgnore, ahk_class CabinetWClass
#IfWinNotActive, ahk_group altScrollIgnore
!Up::Send {WheelUp} ; Scroll window up and down
!Down::Send {WheelDown}
!Left::Send {WheelLeft} ; Scroll window left and right 
!Right::Send {WheelRight}
#IfWinNotActive
