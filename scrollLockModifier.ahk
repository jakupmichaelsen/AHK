#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
IniRead, browser, splashUI.ini, Browser, path
SendMode Input
#NoTrayIcon
TrayTip, %A_ScriptName%, Script loaded
AutoReload()

Ins::
	state :=  GetKeyState("ScrollLock", "T")
	if state
		SetScrollLockState, Off
	if !state 
		SetScrollLockState, On
	ToolTip, ScrollLock toggled 
	SetTimer, RemoveToolTip, 1000
	Return

RemoveToolTip:
SetTimer, RemoveToolTip, Off
ToolTip
Return 


ScrollLock::	
   KeyWait, ScrollLock                      ; wait for Scrolllock to be released
   KeyWait, ScrollLock, D T0.2              ; and pressed again within 0.2 seconds
   if ErrorLevel
      return
   else if (A_PriorKey = "ScrollLock")
      SetScrollLockState, % GetKeyState("ScrollLock","T") ? "Off" : "On"
   return
*ScrollLock:: return                        ; This forces Scrolllock into a modifying key.

#If GetKeyState("ScrollLock", "T")
j::Left
k::Down
l::Right
i::Up
u::PgUp
o::PgDn
h::Home 
n::End
#If 
