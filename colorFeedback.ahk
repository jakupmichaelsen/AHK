#InstallKeybdHook
#SingleInstance, force
AutoReload()

Menu, Tray, Icon, %A_ScriptDir%\icons\highlight.ico
TrayTip, %A_ScriptName%, Script loaded


trying:
try
	global oWord := ComObjActive("Word.Application")
catch
{
	MsgBox,16, %A_ScriptName%, No active MS Word document found.`rExiting.
	ExitApp
}

^Ins::
	state :=  GetKeyState("ScrollLock", "T")
	if state
		SetScrollLockState, Off
	if !state 
		SetScrollLockState, On
	Return


#If GetKeyState("ScrollLock", "T") && WinActive("ahk_class OpusApp") ; if Word is active 

0:: ; Remove all shading 
	oWord := ComObjActive("Word.Application") 
	oWord.Application.Selection.Range.HighlightColorIndex := 0 ; Auto (off)
	oWord.Application.Selection.Shading.BackgroundPatternColor := -16777216 ; Auto (off)
	oWord.Application.Selection.Range.Font.Color := 1
	Return

; ==== Highlight color ==== 
1::shade("red")
2::shade("yellow")
3::shade("green")

5::shade("black")

; ==== Shading ==== 

shade(color){
	oWord := ComObjActive("Word.Application") 
	if color = red 
		oWord.Application.Selection.Shading.BackgroundPatternColor := 255 ; Red 
	if color = yellow
		oWord.Application.Selection.Shading.BackgroundPatternColor := 65535 ; Yellow
	if color = green 
		oWord.Application.Selection.Shading.BackgroundPatternColor := 65280 ; Green 
	if color = black 
	{
		oWord.Application.Selection.Shading.BackgroundPatternColor := 1 ; Black 
		oWord.Application.Selection.Range.Font.Color := 0xFFFFFF
	}
} 

