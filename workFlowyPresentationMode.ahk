#SingleInstance, force
TrayTip, %A_ScriptName%, Script loaded
AutoReload()
Menu, Tray, Icon, %A_ScriptDir%\icons\P.ico
Progress, B0 X%pX% Y%pY% ZH0 CTFF9900 CW221A0F, WF PRESENTATION MODE
; Clicker in presentation mode 
SetTitleMatchMode, 2
WinActivate, WorkFlowy ahk_class Chrome_WidgetWin_1
WinWaitActive, WorkFlowy ahk_class Chrome_WidgetWin_1
Sleep 100
Send ^{Home}


; RButton::Send !{Right}
; LButton::Send !{Left}
Browser_Back::Send {Up}
Browser_Forward::Send {Down}
F3::send ^{Space}


PgDn::
	WinActivate, - WorkFlowy ahk_class Chrome_WidgetWin_1
	WinWaitActive, - WorkFlowy ahk_class Chrome_WidgetWin_1
	Send +!0 
	Return
PgUp::
	WinActivate, - WorkFlowy ahk_class Chrome_WidgetWin_1
	WinWaitActive, - WorkFlowy ahk_class Chrome_WidgetWin_1
	Send +!9
	Return

; .::gosub Suspend  

Ins::
Suspend:
	Suspend, Toggle
	if A_IsSuspended = 0
	{
		Reload
	}
	Else
	{
		Menu, Tray, Icon, %A_ScriptDir%\icons\Blank.ico
		Progress, Hide
	}
	Return
