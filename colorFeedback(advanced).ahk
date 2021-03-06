#InstallKeybdHook
#SingleInstance, force
AutoReload()

Menu, Tray, Icon, %A_ScriptDir%\icons\highlight.ico

TrayTip, %A_ScriptName%, Script loaded
pX := A_ScreenWidth-330
pY := A_ScreenHeight-70
shading = 0 ; Toggle shading and font highlighting 
simple = 0 ; Toggle simple (3 colors) or advanced (6 colors) mode 

try
	global oWord := ComObjActive("Word.Application")
catch
{
	MsgBox,16, %A_ScriptName%, No active MS Word document found.`rExiting.
	ExitApp
}

global docName := oWord.ActiveDocument.Name ; Get file name of current document 

; Ins::
; 	state :=  GetKeyState("ScrollLock", "T")
; 	if state
; 		SetScrollLockState, Off
; 	if !state 
; 		SetScrollLockState, On
; 	Return


#If GetKeyState("Insert", "T") && WinActive("ahk_class OpusApp") ; if Word is active 

0::
	oWord := ComObjActive("Word.Application") 
	oWord.Application.Selection.Range.HighlightColorIndex := 0 ; Auto (off)
	oWord.Application.Selection.Shading.BackgroundPatternColor := -16777216 ; Auto (off)
	Return

; ==== Highlight color ==== 
1::
	if shading = 0
		wHighlight("red")
	Else
		shade("red")
	Return 
2::
	if shading = 0
		wHighlight("yellow")
	Else
		shade("orange")
	Return 
3::
	if shading = 0
		wHighlight("green") ; ("purple")
	Else
		shade("yellow")
	Return 
4::
	if shading = 0
		wHighlight("turquoise")
	Else
		shade("lime green")
	Return 
5::
	if shading = 0
		wHighlight("black")
	Else
		shade("green")
	Return 
6::
	if shading = 0
		wHighlight("grey")
	Else
		shade("")
	Return 

wHighlight(color) {
	oWord := ComObjActive("Word.Application") 
	currentDoc := oWord.ActiveDocument.Name
	if currentDoc <> %docName% ; Check if current document is the same as the previous one
		Gosub, embed_fonts 
	revState =
	if oWord.ActiveDocument.TrackRevisions 
		revState = True
	oWord.ActiveDocument.TrackRevisions := False
	if oWord.Application.Selection.Type = 1
        oWord.Application.Selection.Words(1).Select
	oWord.Selection.Font.Name := "Traveling _Typewriter"
	if color = red
	{
		hColor := 6
		fColor := 0xFFFFFF
	}
	if color = yellow 
	{
		hColor := 7  
		fColor := 0x000000
	}
	if color = purple 
	{
		hColor := 12 
		fColor := 0xFFFFFF
	}
	if color = turquoise
	{
		hColor := 3 
		fColor := 0x000000
	}
	if color = black
	{
		hColor := 1 
		fColor := 0xFFFFFF
	}
	if color = grey 
	{
		hColor := 15 
		fColor := 0x000000
	}
	if color = green 
	{
		hColor := 4
		fColor := 0x000000
	}
	oWord.Application.Selection.Range.HighlightColorIndex := hColor
	oWord.Application.Selection.Range.Font.Color := fColor
	if revState
		oWord.ActiveDocument.TrackRevisions := True
	Send {Right}
	docName := oWord.ActiveDocument.Name ; Update document name 
}

embed_fonts:
	oWord := ComObjActive("Word.Application") 
	oWord.ActiveDocument.EmbedTrueTypeFonts := True
	oWord.ActiveDocument.SaveSubsetFonts := True
	oWord.ActiveDocument.DoNotEmbedSystemFonts := True
	Return 

; ==== Shading ==== 

shade(color){
	oWord := ComObjActive("Word.Application") 
	if color = red 
		oWord.Application.Selection.Shading.BackgroundPatternColor := 255 ; Red 
	if color = yellow
		oWord.Application.Selection.Shading.BackgroundPatternColor := 65535 ; Yellow
	if color = green 
		oWord.Application.Selection.Shading.BackgroundPatternColor := 65280 ; Green 
	if color = orange 
		oWord.Application.Selection.Shading.BackgroundPatternColor := 39423 ; Orange 
	if color = lime green 
		oWord.Application.Selection.Shading.BackgroundPatternColor := 52377 ; Lime green 
} 

