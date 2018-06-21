#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
#NoTrayIcon
SendMode Input
TrayTip, %A_ScriptName%, Script loaded
AutoReload()
pX := A_ScreenWidth-330
pY := A_ScreenHeight-70
paragraph := 0 

SetNumLockState, Off


~CapsLock::
	if (A_PriorHotkey <> "~CapsLock" or A_TimeSincePriorHotkey > 400)
	{
	    ; Too much time between presses, so this isn't a double-press.
	    KeyWait, CapsLock
	    return
	}
	Send, {CapsLock}
	Gosub, toggleNumLock
	Return

toggleNumLock:
	state :=  GetKeyState("NumLock", "T")
	if state
	{
		SetNumLockState, Off
		Progress, Hide
	}
	if !state 
	{
		SetNumLockState, On
		Progress, B0 X%pX% Y%pY% ZH0 CTFF9900 CW221A0F, capsLockModifier,,capsLockModifier.ahk
		WinSet, Transparent, 200, capsLockModifier.ahk
		SetTimer, blink, 60000
	}
	Return

blink:
	loop 2
	{
		if A_Index = 1
		{
			loop 10
			{
				step := 50 + A_Index * 20
				WinSet, Transparent, %step%, capsLockModifier.ahk
				Sleep 100
			}
		}
		Else
		{
			WinSet, Transparent, 200, capsLockModifier.ahk
			SetTimer, blink, Off
		}
	}
	Return 


#If GetKeyState("NumLock", "T")

; a::Left
; s::Down
; d::Right
; w::Up
; q::^Left
; e::^Right
; f::WheelDown
; r::WheelUp
Left::^Left
Right::^Right
Tab::Gosub, nextParagraph
+Tab::Gosub, lastTable


lastTable:
	lastTable := oWord.ActiveDocument.Tables.Count 
	oWord.ActiveDocument.Tables(lastTable).Select
	Return 



nextParagraph:
	if WinActive("ahk_class OpusApp")
	{
		oWord := ComObjActive("Word.Application") 
		pCount := oWord.ActiveDocument.Paragraphs.Count
		thisParagraph := oWord.ActiveDocument.Range(0, oWord.Selection.End).Paragraphs.Count ; Get currenc paragraph #
		paragraph := thisParagraph + 1
		paragraph++
		splashNotify("thisParagraph" thisParagraph "pCount" pCount  "paragraph" paragraph, 2000) 
		if paragraph > %pCount% 
		{
			oWord.ActiveDocument.Paragraphs(1).Range.Select 
			Send, {Left}
		}
		Else
		{
			Try 
			{
				oWord.ActiveDocument.Paragraphs(paragraph).Range.Select 
				Send, {Left}
			}
		}		
	}
	Else
		Send, {Tab}
	Return 

prevParagraph:
	if WinActive("ahk_class OpusApp")
	{
		oWord := ComObjActive("Word.Application") 
		pCount := oWord.ActiveDocument.Paragraphs.Count
		thisParagraph := oWord.ActiveDocument.Range(0, oWord.Selection.End).Paragraphs.Count ; Get currenc paragraph #
		paragraph := thisParagraph + 1
		paragraph--
		splashNotify("thisParagraph" thisParagraph "pCount" pCount  "paragraph" paragraph, 2000) 
		if paragraph <> 1 
		{
			Try 
			{
				oWord.ActiveDocument.Paragraphs(paragraph).Range.Select 
				Send, {Left}
			}
		}
		Else 
		{
			oWord.ActiveDocument.Paragraphs(pCount).Range.Select 
			Send, {Left}
		}
	}
	Else 
		Send +{Tab}
	Return 


#If 
