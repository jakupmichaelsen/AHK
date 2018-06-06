#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
IniRead, browser, splashUI.ini, Browser, path
SendMode Input
Menu, Tray, Icon, D:\Dropbox\Code\AHK\icons\C.ico
TrayTip, %A_ScriptName%, Script loaded
AutoReload()
#IfWinActive, ahk_class Chrome_WidgetWin_1
PgUp::^PgUp
PgDn::^PgDn
Browser_Back::Send {Up}
Browser_Forward::Send {Down}

::..::
	Send {F11}
	Return 