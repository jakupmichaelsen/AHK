#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
IniRead, browser, splashUI.ini, Browser, path
SendMode Input
TrayTip, %A_ScriptName%, Script loaded
AutoReload()

#IfWinActive, ahk_class AcrobatSDIWindow
½::
Browser_Back::
	send, u 
	Sleep 100
	Click, 1105, 90 
	Sleep 1500
	Send {Esc}
	Return 
