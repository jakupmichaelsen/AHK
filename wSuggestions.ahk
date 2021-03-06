#NoTrayIcon
#SingleInstance, force 
#IfWinActive, ahk_class OpusApp 
AutoReload()
tracking = 1
cCount = 1

!BackSpace::Gosub, strikeFont
Tab::Gosub, nextComment
+Tab::Gosub, prevComment
^Tab::Gosub, replaceSelection
Ins::Gosub, myText 
+!^m::Gosub, balloonComment

+!e::ComObjActive("Word.Application").ActiveDocument.Comments(ComObjActive("Word.Application").ActiveDocument.Comments.Count).Edit

balloonComment:
	oWord := ComObjActive("Word.Application")
	oWord.ActiveDocument.TrackRevisions := False 
	oWord.Selection.Font.Color := -16777216
	oWord.Selection.Font.Underline := 0
	oWord.Selection.Font.Bold := 0
	oWord.Selection.Font.StrikeThrough := 0
	oWord.Application.Selection.Range.HighlightColorIndex := 0 ; Auto (off)
	oWord.Application.Selection.Shading.BackgroundPatternColor := -16777216 ; Auto (off)
	Send, 💬
	Send, +{Left}
	Send, !^m
	Send, {ShiftUp}{CtrlUp}{AltUp} ; Unstick 
	Return

nextComment:
	oWord := ComObjActive("Word.Application")
	nComments := oWord.ActiveDocument.Comments.Count
	if ComObjActive("Word.Application").Selection.Information(26) ; wdInCommentPane
	; if nComments < 1 ; If no comments 
	; 	Send, {Tab}
	{
		oWord.ActiveDocument.Comments(cCount).Edit
		cCount++
		if cCount > %nComments%
			Reload
	}
	Else 
	{
		Send, {Tab}
	}
	Return
prevComment:
	oWord := ComObjActive("Word.Application")
	nComments := oWord.ActiveDocument.Comments.Count
	if ComObjActive("Word.Application").Selection.Information(26) ; wdInCommentPane
	; if nComments < 1 ; If no comments 
	; 	Send, {Tab}
	{
		cCount--
		if cCount < 1
			cCount = %nComments%
		oWord.ActiveDocument.Comments(cCount).Edit
	}
	Else
		Send, +{Tab} 
	Return




myText:
	oWord := ComObjActive("Word.Application")
	oWord.Selection.Font.Bold := 1 ; Bold font 
	oWord.Selection.Font.Color := 16748574  ; Blue font 
	Return 

strikeFont:
	oWord := ComObjActive("Word.Application")
	oWord.Selection.Font.Bold := 1 ; Bold font 
	oWord.Selection.Font.Color := -587137152 ; Grey font
	oWord.Selection.Font.StrikeThrough := 1 ; Strikethrough selection
	Return

replaceSelection:
	run %A_ScriptDir%\capsLockModifier.ahk
	oWord := ComObjActive("Word.Application")
	oWord.ActiveDocument.TrackRevisions := False	

	if oWord.Selection.Type = 1 ; If nothing is selected... 
	{
		splashNotify("Selecting word", "typing") 
		oWord.Selection.Words(1).Select ; ...select the current word 
		; Send, +{Left}
		; Sleep 1500
	}
	Else
		splashNotify("Your selection") 
	
	currentFont := oWord.Selection.Font.Name ; Get selection font 
	
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	oWord.Selection.Copy ; Copy selection 
	
	StringRight, rChar, Clipboard, 1 ; Get right character 
	StringLen, wLen, Clipboard ; Get string length 
	
	; Reset comma and period variables 
	isComma = 0 
	isPeriod = 0 
	if Clipboard = ,
		isComma = 1
	if Clipboard = .
		isPeriod = 1
		
	oWord.Selection.Font.Bold := 1 ; Bold font 
	oWord.Selection.Font.Color := -587137152 ; Grey font
	if tracking = 1
	{
		oWord.ActiveDocument.TrackRevisions := True 
		Send, {Delete}
		oWord.ActiveDocument.TrackRevisions := False	
		Send, {Right}{Left 2}
	}
	Send, {Right}
	Sleep 200
	if (isComma or isPeriod) 
		oWord.Selection.Font.Subscript := 1 ; Use subscript if selection is comma or period 
	; if (wLen != 1 && rChar != A_Space) ; If selection isn't a single space: insert space
	; 	Send, {Space}
	Gosub, myText
	Send, → ; Send arrow 
	oWord.Selection.Font.Subscript := 0 ; Subscript off 
	oWord.Selection.Font.Name := currentFont ; Change Word's fallback font to currentFont
	
	if wLen != 1
	{ ; Unless string is 1 character, send: 
		Send, {Space}
		Send, {Left}
	}
	
	Clipboard = %ClipSaved%  ; Restore Clipboard
Return 
