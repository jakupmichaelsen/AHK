#SingleInstance, force
#Hotstring, * ?
splashNotify(A_ScriptName " reloaded") 
SendMode Input
AutoReload()

paragraph := 0 
SetScrollLockState, Off



#If GetKeyState("ScrollLock", "T")

j::Send, {left}
l::Send, {right}
i::Send, {up}
k::Send, {down}
+j::Send, +{Left}
+l::Send, +{Right}
+i::Send, +{Up}
+k::Send, +{Down}
^j::Send, ^{Left}
^l::Send, ^{Right}
^i::Send, ^{Up}
^k::Send, ^{Down}
+^j::Send, +^{Left}
+^l::Send, +^{Right}
+^i::Send, +^{Up}
+^k::Send, +^{Down}

#If 




/*
LCtrl & j::Gosub, controlLeft
LCtrl & l::Gosub, controlRight
LCtrl & i::Gosub, controlUp
LCtrl & k::Gosub, controlDown
j::Gosub, leftKey
l::Gosub, rightKey
i::Gosub, upKey
k::Gosub, downKey
+j::Gosub, shiftLeftKey
+l::Gosub, shiftRightKey
+i::Gosub, shiftUpKey
+k::Gosub, shiftDownKey

controlLeft:
	SetKeyDelay, 185,
	if GetKeyState("Shift", "P")
		Send, +^{Left}
	Else
		Send, ^{Left}
	Return

controlRight:
	SetKeyDelay, 185,
	if GetKeyState("Shift", "P")
		Send, +^{Right}
	Else
		Send, ^{Right}
	Return

controlUp:
	SetKeyDelay, 185,
	if GetKeyState("Shift", "P")
		Send, +{Up}
	Else
		Send, ^{Up}
	Return

controlDown:
	SetKeyDelay, 185,
	if GetKeyState("Shift", "P")
		Send, +{Down}
	Else
		Send, ^{Down}
	Return


leftKey:
	Send, {Left}
	Return

rightKey:
	Send, {Right}
	Return

upKey:
	Send, {Up}
	Return

downKey:
	Send, {Down}
	Return

shiftLeftKey:
	Send, +{Left}
	Return
shiftRightKey:
	Send, +{Right}
	Return
shiftUpKey:
	Send, +{Up}
	Return
shiftDownKey:
	Send, +{Down}
	Return
*/

~CapsLock::
	if (A_PriorHotkey <> "~CapsLock" or A_TimeSincePriorHotkey > 400)
	{
	    ; Too much time between presses, so this isn't a double-press.
	    KeyWait, CapsLock
	    return
	}
	Send, {CapsLock}
	Gosub, toggleScrollLock
	Return

toggleScrollLock:
	if GetKeyState("ScrollLock", "T")
	{
		SetScrollLockState, Off
		Progress, hide 
	}
	Else 
	{
		SetScrollLockState, On
		splashNotify("capsLockModifier", "Tray", 600000) 
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


