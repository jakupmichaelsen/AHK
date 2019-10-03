#NoEnv ; Significantly improves performance by not looking up empty variables as environmental variables
AutoReload()

SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#If GetKeyState("NumLock", "T") 
	w::Send, {Up}
	a::Send, {Left}
	s::Send, {Down}
	d::Send, {Right}
	
	^w::Send, ^{Up}
	^a::Send, ^{Left}
	^s::Send, ^{Down}
	^d::Send, ^{Right}
	
	+w::Send, +{Up}
	+a::Send, +{Left}
	+s::Send, +{Down}
	+d::Send, +{Right}
	
	+^w::Send, +^{Up}
	+^a::Send, +^{Left}
	+^s::Send, +^{Down}
	+^d::Send, +^{Right}
#If 

~CapsLock::
	if (A_PriorHotkey <> "~CapsLock" or A_TimeSincePriorHotkey > 400)
	{
	    ; Too much time between presses, so this isn't a double-press.
	    KeyWait, CapsLock
	    Return
	}
	Send, {CapsLock}
	Gosub, toggleNumLock
	Return

toggleNumLock:
	if GetKeyState("NumLock", "T")
		Gosub, singleKeysOff
	Else 
		Gosub, singleKeysOn
	Return 

singleKeysOff:
	SetNumLockState, Off
	Progress, hide 
	Return 

singleKeysOn:
	SetNumLockState, On
	splashNotify("ASDW navigation", "statusBar", 600000) 
	Return

splashNotify(text, position="top", timeout=2000, fontSize= 14, transparency=200){ 
	global
	iniRead()
	if timeout <> 0
		SetTimer, notifyTimeout, %timeout%
	if position = top
	{
		y := 0
		x := 0
	}	
	if position = middle
	{
		y := A_ScreenHeight / 2 - 70
		x := 0
	}	
	if position = bottom 
	{
		y := A_ScreenHeight - 70
		x := 0
	}	
	if position = typing
	{
		y := A_ScreenHeight / 4
		x := A_ScreenWidth / 2 - 150 ; Default Progress width is 300
	}
	if position = tray
	{
		y := A_ScreenHeight / 1.085
		x := A_ScreenWidth / 1.23
	}
	if position = statusbar
	{
		y := A_ScreenHeight / 1.085
		x := 0
	}
	if position <> statusbar 
		Progress, B0 X%x% Y%y% ZH0 fs%fontSize% CTFF9900 CW221A0F, %text%, , splashNotify.ahk, %font%
	else 
		Progress, B0 X%x% Y%y% ZH0 fs%fontSize% W%A_ScreenWidth% CWFF9900 CT221A0F, %text%, , splashNotify.ahk, %font%
	WinSet, Transparent, %transparency%, splashNotify.ahk
}

notifyTimeout:
	Progress, Off
	Return 
