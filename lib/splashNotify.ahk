splashNotify(text, timeout=0, position="top"){ ; Pass text to display splash notification. Timeout defaults to off. Position to top. 
	global
	IniRead()
	if timeout <> 0
		SetTimer, notifyTimeout, %timeout%
	if position = top
		y := 0
	if position = center
		y := A_ScreenHeight / 2 - 70
	if position = bottom 
		y := A_ScreenHeight - 70
	if position = typing
		y := A_ScreenHeight / 1.7
	if position <> typing
		Progress, B0 X0 Y%y% ZH0 CTFF9900 CW221A0F, %text%, , splashNotify.ahk, 
	Else
		Progress, B0 Y%y% ZH0 CTFF9900 CW221A0F, %text%, , splashNotify.ahk, 
	WinSet, Transparent, 200, splashNotify.ahk
}

notifyTimeout:
	Progress, Off
	Return 
