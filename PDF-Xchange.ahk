#SingleInstance, force 
IniRead, ahk, splashUI.ini, paths, ahk

SplashImageGUI(Picture, X, Y, Duration, Transparent = true)
{
Gui, XPT99:Margin , 0, 0
Gui, XPT99:Add, Picture,, %Picture%
Gui, XPT99:Color, white
Gui, XPT99:+LastFound -Caption +AlwaysOnTop +ToolWindow -Border
If Transparent
{
Winset, TransColor, white
}
Gui, XPT99:Show, x%X% y%Y% NoActivate
; SetTimer, DestroySplashGUI, -%Duration%
return

; DestroySplashGUI:
; Gui, XPT99:Destroy
; return
}
 
splashUI("r", PDF-XChange|Word, "PDF-XChange or Word?")
if choice <>
{
	if choice = 1
		SplashImage1 = %ahk%\images\PDF-XChange.png
	if choice = 2
		SplashImage2 = %ahk%\images\Word.png
}

Ins::
	Toggle++
	if  mod(Toggle, 2) = 1 
		SplashImageGUI(SplashImage, "Center", "Center", 2000, true)
	Else 
		Gui, XPT99:Destroy
	Return


+Ins::
	run rundll32.exe shell32.dll Options_RunDLL 1
	WinWaitActive Egenskaber for Proceslinje og menuen Start 
	Send, s{Enter}
	Return

^Ins::ExitApp
