#NoEnv ; Significantly improves performance by not looking up empty variables as environmental variables 
#SingleInstance, force

IniRead, onedrive, splashUI.ini, paths, onedrive
Menu, Tray, Icon, %A_ScriptDir%\icons\word-128.ico
TrayTip, %A_ScriptName%, Script loaded
AutoReload()

global onedrive 
global toggle = 0
global array := { "Genitiv": "Genitive","Præposition før that-sætning": "Preposition with that-conjunction","Adjektiv-adverbium": "adjectives & adverb","Abstrakte substantiver": "abstract nouns","Adverbiers placering": "adverb placemen","Forkert ordvalg": "word choice","Danisme": "Danishism","Spring i tid": "Switching verb time","Gradbøjning af adjektiver": "adjective form","Kongruensfejl": "S-V agreement","Forkert ordklasse": "part of speech","Formuleringsfejl": "wording","Pronominer": "pronouns","Ordstilling": "word order","Forkert præposition": "wrong preposition","Proprier": "proper nouns","Substantivers flertal": "plural nouns","Talesprog": "vernacular","Udvidet tid": "progressive ","Verbets tider": "verb form","Syntaks": "syntax","Tegnsætning": "punctuation" }

global sentenceTable = 1
global commaTable = 1
global copy2XL = 0
global includeErrorType = 1
global asciiBarChart = 1
global screenUpdating = 0
global markErrorInline = 0
global errorCounter = 0

global xlSnippetsVisible = 1
global xlPasteFormatting = 1 
global feedbackTable = 1
global inLineComments = 0
global copyWholeSentence = 1
global ttsErrors = 0
global feedbackInsert = "top" ; "bottom"

; === Files === 

global xlSnippets
global $comment
global commentsTemplate
global feedbackTemplate

xlSnippets = %A_ScriptDir%\xlSnippets.xlsx
$comment = %onedrive%\Adjunkt\Grading\Templates\§ comment.rtf
commentsTemplate = %onedrive%\Adjunkt\Grading\Templates\comment.rtf ; "comments_focus.rtf"
feedbackTemplate = %onedrive%\Adjunkt\Grading\exam_DigiA.rtf ; "language.rtf" ; "exam.rtf" ; "language_sans_comma.rtf" ; "language.rtf" ; "comment_interpretation.rtf" ; "language_terms_interpretation_wording_word-choice.rtf" ; "grammar.rtf" "language_terms_interpretation.rtf"

if copy2XL = 1  
	TrayTip, %A_ScriptName%, NB: Copying to Excel is active, 4, 17


#IfWinActive, ahk_class OpusApp
:*:"""::Dine grammatiske fejl er markeret med rød – se om du finde hvad er hvad ud fra nedenstående tabel. ; General comment
:*:\comment::
	Sleep 500
	oWord := ComObjActive("Word.Application").Selection.InsertFile(commentsTemplate)
	Send {Left}
	Return 
+^'::feedback("Genitiv")
+^-::feedback("Sammensat adjektiv")
+^+::feedback("Præposition før that-sætning")
+^a::feedback("Adjektiv-adverbium")
+^b::feedback("Abstrakte substantiver") ; ("Fejl ved artikler")
+^c::
	if feedbackTable = 1
		feedbackTable("Spelling") 
	if inLineComments = 1
		feedbackNote("Stavefejl")
	Return 
+^d::feedback("Adverbiers placering")
+^e::feedback("Fejl ved artikler")
+^f::feedbackNote("Overflødigt")
!^g:: 
	if feedbackTable = 1
 		feedback("Word choice") 
	if inLineComments = 1
		feedbackNote("Forkert ordvalg")
	Return 
+^h::Gosub, blackHighlighter
!^i::feedback("Danisme")
; +^i::feedback("Spring i tid")
+^j::feedback("Gradbøjning af adjektiver")
!^k::feedback("Kongruensfejl")
!^l::feedback("Forkert ordklasse")
!^m::feedback("Formuleringsfejl")
!^n::feedback("Pronominer")
!^o::feedback("Ordstilling")
!^p::feedback("Forkert præposition")
+^r::feedback("Proprier")
+^s::feedback("Substantivers flertal")
+^t::feedbackNote("Informal")
!^u::feedback("Udvidet tid")
!^v::feedback("Verbets tider")
+^x::feedback("Fejl ved hjælpeverber")
+^y::feedback("Syntaks")
+^,::
	if commaTable = 1
		feedbackTable("Comma")
	Else
		Gosub, punctuationError
	Return 
+^.::Gosub, insertPeriod ; punctuationError
^!"::Gosub, wrapQuotationMarks
^!,::Gosub, insertComma
^!.::Gosub, insertPeriod ; movePeriod
^!x::Gosub, cutAndMove
; +^Space:: Gosub, scrollIntoView ; Gosub, errorFont
+^Enter::Gosub, removeNewLine
+Ins::Gosub, singleUppercase


+^0::Gosub, clearFormatting 
+BackSpace::
	Gosub, greyFont  
	Gosub, boldFont
	Return 
!BackSpace::Gosub, strikeFont
+!BackSpace::
	Gosub, courier 
	Gosub, greyFont
	Gosub, boldFont
	Return 
+^F2::Gosub, toggleFeedbackTable ; toggleB1B2
^F2::Gosub, closePane ; removeSplit

F2::Gosub, togglePanes ; Split window and toggle panes 
; ^1::
; 	Gosub, saveB1
; 	splashNotify("B1 saved", "bottom")
; 	Return
!F3::
	; "Select All Text With Similar Formatting" 
	; Requires 'Select' to be in the position 1 in the Quick Access Toolbar
	Send !1s
	Sleep 500
	Gosub, wStringInfo
	splashNotify("Len: " . wLen) 
	Return 
+^F3::Gosub, saveB2
^F3::Gosub, saveB1
F3::Gosub, toggleB1B2
	; Gosub, saveB1
	; Gosub, gotoB2
	; splashNotify("B2", "bottom")
	; Gosub, scrollIntoView
	Return 
; ^2::
; 	Gosub, saveB2
; 	splashNotify("B2 saved", "bottom")
; 	toggle = 1 ; Reset toggle counter for toggling between bookmarks 
; 	Return
F4::Gosub, inlineComment
+F4::Gosub, tableComment 

F6::
	Gosub, inlineComment
	Gosub, xlSnippets
	Return 
F7::Gosub, switchWords
F8::Gosub, toggleFeedbackTable ; checkOrInsertTable
+F8::Gosub, generalComments
^space::feedbackTable("interpretation")
^½::wordToTable("terminology")

; !Space::
; splashList("", "interpretation|terminology",1,30)
; #fff todo → +F6::feedbackNote("Good topic sentence")
; #fff todo → +F7::feedbackNote("Good thesis statement")
; 	if choice = terminology
; 		wordToTable("terminology")
; 	if choice = interpretation 
; 		feedbackTable("interpretation")
; 	Return 

#If GetKeyState("CapsLock", "T") and WinActive("ahk_class OpusApp")
Tab::Gosub, togglePanes ; Split window and toggle panes 
+Tab::Gosub, lastTable

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
#If  

feedback(what) {
	if feedbackTable = 1
		feedbackTable(what)
	Else
		feedbackNote(what)
}

feedbackTable(what) {
	oWord := ComObjActive("Word.Application")
	oWord.ActiveDocument.Bookmarks.Add("progress") ; Set progress bookmark

	Gosub, saveProgress 
	Gosub, checkOrInsertTable
	Gosub, gotoProgress

	if what = Formuleringsfejl
		what = Wording
	if what = Forkert ordvalg
		what = Word choice

	if (what = "Wording" || what = "Word choice" || what = "Spelling" || what = "Comma" || what = "interpretation") ; Feedback type  
		whichWhat = comment 
	Else
		whichWhat = grammar

	tag := what ; Copy what to tag variable 
	StringReplace, tag, tag, %A_Space%, _, All ; Replace space in tag with underscore 
	StringReplace, tag, tag, -, _, All ; Replace dash in tag with underscore 
	barMark := "mark_" . tag 
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 

	if what <> interpretation
		Gosub, selectionOrWord ; Check if something's selected, otherwise select current/closest word 
	if what = interpretation 
	{
		if oWord.Selection.Type = 1 ; If no selection 
			Gosub, copyWholeSentence
		Else 
			oWord.Selection.Copy ; Copy selection 
		; Clipboard = %Clipboard% ; Remove formatting 
		Gosub, saveProgress
	}

	if what <> comma ; Format selection 
	{
		Gosub, underlineFont
		Gosub, boldFont
		; sleep 100
	}
	if screenUpdating = 0
		oWord.Application.ScreenUpdating := False

	if what <> interpretation ; interpretation has already been copied 
	{
		if copyWholeSentence = 1
			Gosub, copyWholeSentence ; Copy current sentence 
		Else 
			Gosub, copyContext ; Copy context 
	}
	if what = comma 
		Clipboard = %Clipboard% ; Remove formatting 
	Gosub, gotoProgress
	if markErrorInline = 0 
		Gosub, clearFormatting
	splashNotify(what, "typing", 1500, 14, 200)
 	if whichWhat = comment
		pasteToTable(tag)
	if whichWhat = grammar 
	{ ; TODO: Speed up xlPaste as well !! <--
		if copy2XL = 1
			xlPaste(Clipboard, tag)
		if sentenceTable = 1 
		{
			if includeErrorType = 1
				pasteToTable("sentence", what)
			Else 
				pasteToTable("sentence")
		}
		if asciiBarChart = 1
			errorTracker(barMark, what)
	}

	Clipboard = %ClipSaved%  ; Restore clipboard
	Gosub, gotoProgress
	Gosub, scrollIntoView
 	Send, {CtrlUp}{Right}
	Gosub, saveDoc 
	oWord.Application.ScreenUpdating := True
}

errorTracker(barMark, what) { 
	oWord := ComObjActive("Word.Application") 
 	if oWord.ActiveDocument.Bookmarks.Exists(barMark) ; If bookmark exists 
		oWord.ActiveDocument.Bookmarks(barMark).Select ; Select bookmark in Word
 	Else ; If bookmark doesn't exist
 	{	
		oWord.ActiveDocument.Bookmarks("grammar").Select ; Select bookmark in Word
		oWord.Selection.InsertRowsAbove
		oWord.Selection.ParagraphFormat.SpaceBefore := 0
		oWord.Selection.ParagraphFormat.SpaceAfter := 0
		oWord.Selection.TypeText(what . ":")
		oWord.Selection.Cells(1).Next.Select 
		oWord.Selection.HomeKey
	}
	oWord.Selection.TypeText("■")
	oWord.ActiveDocument.Bookmarks.Add(barMark) ; Add/update bar chart bookmark 
	if errorCounter = 1
		Gosub, updateErrorCounter ; <- speed up <-- TODO!
	if sentenceTable <> 1 ; Only update table progress if not using sentence table 
		Gosub, saveTableProgress
}

pasteToTable(bookmark, text="") {
	oWord := ComObjActive("Word.Application") 
	Gosub, checkOrInsertTable
	oWord.ActiveDocument.Bookmarks(bookmark).Select 
	; oWord.Selection.LanguageID := 1024 ; No proofing
	if text <> ; If text not empty  
	{
		oWord.Selection.HomeKey
    	oWord.Selection.Font.Bold := True 
    	oWord.Selection.TypeText(text . ":  ")
    	oWord.Selection.Font.Bold := False
	}
	StringLen, wLen, Clipboard 
	if (wLen < 15 and bookmark = "interpretation")
	{
		oWord.Selection.ParagraphFormat.Alignment := 2 ; wdAlignParagraphRight = 2
		oWord.Selection.PasteAndFormat(20) ; Paste with destination formatting 
		oWord.Selection.TypeParagraph ; Paragraph break 
		oWord.Selection.ParagraphFormat.Alignment := 0 ; wdAlignParagraphLeft = 0
	}
	Else
	{	
		oWord.Selection.PasteAndFormat(20) ; Paste with destination formatting 
		oWord.Selection.TypeParagraph ; Paragraph break 
	}
	oWord.ActiveDocument.Bookmarks.Add(bookmark)
	Gosub, saveTableProgress
	text = ; Clear text variable 
}

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

blackHighlighter:
	Gosub, selectionOrWord
	oWord := ComObjActive("Word.Application")
	oWord.Application.Selection.Range.HighlightColorIndex := 1 ; Black 
	oWord.Application.Selection.Range.Font.Color := 0xFFFFFF ; White 
	Send {Right}
	Return 

updateErrorCounter:
	oWord := ComObjActive("Word.Application") 
	oWord.ActiveDocument.Bookmarks("error_counter").Select ; Select bookmark in Word
	Sleep 200
	oWord.Selection.Words(1).Select ; Select the current word ; Send, +{Right}
	Sleep 200
	oWord.Selection.Copy
	Clipboard++
	oWord.Selection.Paste 
	Send, {Space}
	Return 

closePane: 
	oWord := ComObjActive("Word.Application")
	Try 
		oWord.ActiveWindow.ActivePane.Close
	Return 

removeSplit:
	oWord := ComObjActive("Word.Application")
	oWord.ActiveWindow.SplitVertical := 100 ; Remove window split 
	oWord.ActiveWindow.View.Type := 3 
	Return 

setupPanes:
	aW := ComObjActive("Word.Application").ActiveWindow
	if aW.Panes.Count > 2 ; Ensure there aren't more than two panes 
		aW.Panes(3).Close 
	Sleep 100 
	if aW.Panes.Count < 2 ; If only 1 pane, split vertically at 85% and setup the panes
	{	
		aW.SplitVertical := 85
		Sleep 100 
		aW.Panes(1).View.Type := 3 ; wdPrintView = 3
 		aW.Panes(2).View.Type := 6 ; wdWebView = 6
 		aW.Panes(2).View.Zoom.Percentage := 140
	}
	Return 

togglePanes:
	aW := ComObjActive("Word.Application").ActiveWindow

	; Get current pane index 
	pIndex := aW.ActivePane.Index

	; Setup bookmarks depending on current pane: 
	if pIndex = 1 
		Gosub, saveProgress
	if pIndex <> 1 
		Gosub, savePaneProgress

	; Check and/or setup window split:
	if aW.SplitVertical <> 85 
		aW.SplitVertical := 85

	; Activate next pane: 
	aW.ActivePane.Next.Activate

	; Get new pane index:
	pIndex := aW.ActivePane.Index 

	; splashNotify(pIndex, "top", 22, 200) 
	
	; ; Format other pane: 
	if pIndex <> 1 
 	{
 		; Setup if not already formatted: 
 		if aW.Panes(pIndex).View.Type <> 6 ; wdWebView = 6
 			aW.Panes(pIndex).View.Type := 6 
 		if aW.Panes(pIndex).View.Zoom.Percentage <> 140
 			aW.Panes(pIndex).View.Zoom.Percentage := 140
 		Gosub, gotoPaneProgress
	}

	; Select bookmark in new pane: 
	if pIndex = 1
		Gosub, gotoProgress
	if pIndex <> 1 
 		Gosub, gotoPaneProgress
	; Save document 
	Gosub, saveDoc
	
	; Blink cursor: 
	; Sleep, 100
	; Send, {Right}
	; Sleep, 25
	; Send, {Left}

	Return 

saveDoc:
	ComObjActive("Word.Application").ActiveDocument.Save
	Return 



punctuationError:
	Gosub, inlineComment
	Send, {Right}
	Send, +{Left 2}
	oWord.Selection.Font.Subscript := 1
	Send, {Right}{Left}←%A_Space%
	Return

selectSentence:
	oWord := ComObjActive("Word.Application")
	oWord.Selection.MoveLeft(Unit:=3)  ; Move to start of sentence
	oWord.Selection.ExtendMode := True ; Turns on extend mode 
	oWord.Selection.MoveRight(Unit:=3) ; Expand selection to whole sentence
	oWord.Selection.ExtendMode := False ; Turns off extend mode 
	Return 

; selectionToTable(bookmark){
; 	oWord := ComObjActive("Word.Application") 
; 	if screenUpdating = 0
; 		oWord.Application.ScreenUpdating := False 
; 	toggle = 0 ; Reset toggle counter 
; 	BlockInput, On
; 	Gosub, selectionOrWord
; 	ClipSaved = %ClipboardAll%  ; Save clipboard
; 	Clipboard = ; Clear clipboard 
; 	if bookmark <> "interpretation"
; 	{
; 		Gosub, boldFont
; 		Gosub, underlineFont
; 		Gosub, greyFont
; 		Gosub, copySentenceOrContext
; 	}
; 	Sleep 100
; 	oWord.Selection.Copy 
; 	if bookmark = "interpretation"
; 		Clipboard = %Clipboard%
; 	Gosub, saveProgress
; 	Gosub, checkOrInsertTable
; 	oWord.ActiveDocument.Bookmarks(bookmark).Select 
; 	Sleep 100
	; 
; 	if bookmark = interpretation
; 	{
; 		Clipboard = ”%Clipboard%”
; 		oWord.Selection.PasteAndFormat(22) ; Paste without formatting  
; 	}
; 	Else
; 		oWord.Selection.PasteAndFormat(20) ; Paste with destination formatting 
; 	Send, {Enter}
; 	; oWord.ActiveDocument.Range.ParagraphFormat.SpaceBefore := 6
; 	Clipboard = %ClipSaved%  ; Restore clipboard
; 	Sleep 250
; 	oWord.ActiveDocument.Bookmarks.Add(bookmark) ; Update bookmark
; 	Gosub, saveTableProgress
; 	Gosub, gotoProgress
; 	; Gosub, scrollIntoView
; 	Send, {Right}
; 	oWord.Application.ScreenUpdating := True
; }

wordToTable(bookmark){
	oWord := ComObjActive("Word.Application") 
	if screenUpdating = 0
		oWord.Application.ScreenUpdating := False
	toggle = 0 ; Reset toggle counter 
	Gosub, selectionOrWord
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	oWord.Selection.Copy 
	splashNotify(bookmark . ": " . Clipboard, "typing")
	Gosub, saveProgress
	; Sleep 100
	Gosub, checkOrInsertTable
	Try 
		oWord.ActiveDocument.Bookmarks(bookmark).Select ; Select bookmark (if it exists)
	Catch {
		MsgBox,16, %A_ScriptName%,Bookmark `"%bookmark%`" not found.`r`rTry again.
		Reload
	}
	oWord.Selection.PasteAndFormat(22) ; Paste without formatting  
	oWord.Selection.TypeText(", ")
	oWord.ActiveDocument.Bookmarks.Add(bookmark) ; Update terminology bookmark

	Gosub, saveTableProgress
	Clipboard = %ClipSaved%  ; Restore clipboard
	Gosub, gotoProgress
	Gosub, scrollIntoView
 	Send, {CtrlUp}{Right}
	Gosub, saveDoc 
	oWord.Application.ScreenUpdating := True
}

toggleB1B2:
	toggle++
	if  mod(toggle, 2) = 0
	{
		Gosub, gotoB1
		splashNotify("B1") 
	}
	if  mod(toggle, 2) = 1
	{	
		Gosub, gotoB2 
		splashNotify("B2") 
	}
	Return


toggleFeedbackTable:
	Gosub, checkOrInsertTable
	toggle++
	if  mod(toggle, 2) = 0
	{
		Gosub, gotoProgress
		Gosub, scrollIntoView
	}
	if  mod(toggle, 2) = 1
		Gosub, gotoTableProgress
	Return 

tableComment: ; Table paragraph comment - cuts selection and pastes it into table from RTF file 
	oWord := ComObjActive("Word.Application") 
 	oWord.ActiveDocument.TrackRevisions := False
 	ClipSaved = %ClipboardAll%  ; Save clipboard
 	Clipboard = ; Clear clipboard 
	Send ^x 
	; Insert template 
	oWord.Selection.InsertFile(§comment)
	; Paste selection  
    oWord.Application.ActiveDocument.Bookmarks("paste_here").Select ; Select bookmark in Word
	oWord.Selection.PasteAndFormat(16) ; Paste selection, maintaining formatting
	oWord.Application.ActiveDocument.Bookmarks("paste_here").Delete
	Send, {BackSpace}
    oWord.Application.ActiveDocument.Bookmarks("paragraph_comment").Select ; Select bookmark in Word
    oWord.Application.ActiveDocument.Bookmarks("paragraph_comment").Delete
    Clipboard = %ClipSaved%  ; Restore clipboard
    Return

lastTable:
	oDoc := ComObjActive("Word.Application").ActiveDocument
	lastTable := oDoc.Tables.Count 
	oDoc.Tables(lastTable).Select
	Return 

removeNewLine:
	Send, {Del}
	Send, ⎶{Space}{Left}
	feedbackNote("Removed paragraph break")
	Return 

switchWords:
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Send, ^c
	ClipWait
	Gosub, greyFont  
	StringSplit, words, Clipboard, %A_Space%
	Send, {Right}
	Gosub, blueFont
	Gosub, boldFont
	Send, →  
	Send, % A_Space words2 A_Space words1 A_Space
	Clipboard = %ClipSaved%  ; Restore clipboard
	Return 

wrapQuotationMarks:
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Send, ^x 
	ClipWait
	Send, %A_Space%”%Clipboard%”
	Clipboard = %ClipSaved%  ; Restore clipboard
	Return 

colorThenColor(color1, color2){
	oWord := ComObjActive("Word.Application") 
	currentFont := oWord.Selection.Font.Name ; Get selection font 
	if oWord.Selection.Type = 1
		oWord.Selection.Words(1).Select ; Or select the current word first
	Gosub, wStringInfo
	; if wLen = 1
	Gosub boldFont
	Sleep 200
	Gosub, %color1% 
	Send, {Right}
	Sleep 200
	if isComma 
		oWord.Selection.Font.Subscript := 1
	if (wLen != 1 && rChar != A_Space)
		Send, {Space}
	Gosub, boldFont
	Gosub, %color2% 
	Send, →
	oWord.Selection.Font.Subscript := 0
	oWord.Selection.Font.Name := currentFont ; Change Word's fallback font to currentFont
	if wLen != 1 ; if rChar = %A_Space%
	{
		Send, {Space 2}
		Send, {Left}
	}
	Return 
}

feedbackNote(note){
	Gosub, selectionOrWord
	Gosub, inlineComment
	Sleep 100
	Send %note%
	Send, {Right 2}
	Sleep 1000
}


checkOrInsertTable:
 	oWord := ComObjActive("Word.Application") 
 	; All feedback tables have "table_progress" bookmark. If bookmarks doesn't exist, insert feedback table. 
 	if not oWord.ActiveDocument.Bookmarks.Exists("table_progress") 
 	{	
		Gosub, saveProgress 
 		insertFeedbackTemplate(feedbackTemplate)
		Sleep 200
		Gosub, gotoProgress
	}
	Return 

trimNewLine:
	StringReplace, Clipboard, Clipboard, `r,, All
	StringReplace, Clipboard, Clipboard, `n,, All
	Return

copyContext:
	oWord := ComObjActive("Word.Application") 
	wCount := oWord.Selection.Words.Count
	if wCount < 4
	{
		oWord.Selection.MoveLeft(Unit:=2, Count:=5, Extend:=0)
		if wCount = 1
			rCount := 10
		Else 
			rCount := 10 + wCount 
		oWord.Selection.MoveRight(Unit:=2, Count:=rCount, Extend:=1)
	}
	oWord.Selection.Copy
	oWord.Selection.ExtendMode := False
	; Sleep 100
    Return 

RemoveToolTip:
	SetTimer, RemoveToolTip, Off
	ToolTip
	Return 

openMistakesXl: 
	fileName := ComObjActive("Word.Application").ActiveDocument.Path . "\assignmentData.xlsx"
	ComObjActive("Word.Application").Application.ScreenUpdating := True 
	Try 
		xl := ComObjActive("Excel.Application") 
	Catch 
		xl := ComObjCreate("Excel.Application") 
	splashNotify("Trying to open 'assignmentData.xlsx'", "typing") 
	Sleep 500
	Try 
	{
		xl.Workbooks.Open(fileName)
		xl.Visible := True 
	}
	Catch
	{
		splashNotify("'assignmentData.xlsx' not found. Setting up new file.", "typing") 
		Gosub, createMistakesXl
	}
	WinActivate, ahk_class OpusApp
	WinWaitActive, ahk_class OpusApp
	Gosub, gotoProgress
	Sleep 500
	splashNotify("'assignmentData.xlsx' opened. Script will now reload", "typing")
	Sleep 2000
	Reload 
	Return 

createMistakesXl:
	xl := ComObjCreate("Excel.Application") 
	xl.Visible := True 
	xl.Workbooks.Add
	fileName := ComObjActive("Word.Application").ActiveDocument.Path . "\assignmentData.xlsx"
	xl.ActiveSheet.Range("A" . 1).Value := "Filename"
	xl.ActiveSheet.Range("B" . 1).Value := "Sentence"
	xl.ActiveSheet.Range("C" . 1).Value := "Tag" 
	xl.ActiveSheet.Range("D" . 1).Value := "Timestamp"
	xl.Columns("A").ColumnWidth := 30
	xl.Columns("B").ColumnWidth := 140
	xl.Columns("C").ColumnWidth := 20
	xl.Columns("D").ColumnWidth := 14
	xl.Columns("D").NumberFormat := 0
	xl.Rows("1").Font.Bold := True
	xl.ActiveWorkbook.SaveAs(fileName)
	xl.Application.Quit
	Sleep 2000
	xl := 
	; Gosub, openMistakesXl
	splashNotify("'assignmentData.xlsx' created. Script will now reload", "typing")
	Sleep 2000 
	Reload 
	Return 

xlPaste(text, tag){
	Try
	{
		xl := ComObjActive("Excel.Application") 
		; xl := Excel_Get() 
		xl.Workbooks("assignmentData.xlsx").Activate
	}
	Catch 
	{
		Gosub, openMistakesXl
	}

	xl := ComObjActive("Excel.Application") 
	xl.Workbooks("assignmentData.xlsx").Activate
	oWord := ComObjActive("Word.Application") 
	row := xl.ActiveSheet.UsedRange.Rows.Count + 1
	; tag := "#" . tag 
	xl.ActiveSheet.Range("A" . row).Value := oWord.ActiveDocument.Name
	if xlPasteFormatting = 1 
	{
		Try
			xl.Range("B" . row).PasteSpecial(Format := "HTML") ; Formatting might be off, if trimNewLine has been called
		Catch
			xl.ActiveSheet.Range("B" . row).Value := Clipboard ; - in that case paste without formatting 
	}
	Else 
		xl.ActiveSheet.Range("B" . row).Value := Clipboard
	xl.ActiveSheet.Range("C" . row).Value := tag
	xl.ActiveSheet.Range("D" . row).Value := A_Now
	xl.ActiveWorkbook.Save
}


insertCommentsTemplate(name){
	oWord := ComObjActive("Word.Application") 
	Gosub, saveProgress ; Set progress bookmark
 	toggle = 1 ; Reset toggle counter for toggling between bookmarks 
 	Send ^{Home}
 	Sleep 200
 	Send {Enter}
	oWord.Selection.InsertFile(commentsTemplate)
}

insertFeedbackTemplate(name){
	oWord := ComObjActive("Word.Application") 
 	toggle = 1 ; Reset toggle counter for toggling between bookmarks 
 	if feedbackInsert = bottom
 	{
 		Send ^{End}
 		Sleep 200
 		Send ^{Enter}
 	}
 	if feedbackInsert = top
 	{
 		Send ^{Home}
 		Sleep 200
 		Send ^{Enter}
 		Send ^{Home}
 	}
 	else 
 	{	
 		Msgbox Set %feedbackInsert% to either "top" or "buttom". Script will now reload!
 		Reload
 	}
	oWord.Selection.InsertFile(feedbackTemplate)
}

:*:\\::→


copyWholeSentence:
	oWord := ComObjActive("Word.Application") 
	Gosub, saveProgress ; Set progress bookmark
	oWord.Selection.MoveLeft(Unit:=3)  ; Move to start of sentence
	oWord.Selection.ExtendMode := True ; Turns on extend mode
	oWord.Selection.MoveRight(Unit:=3) ; Expand selection to whole sentence
	oWord.Selection.ExtendMode := False ; Turns off extend mode
	oWord.Selection.Copy 
	; Sleep 100
	Return 



generalComments: 
 	; General comment at the top of the document
 	if ComObjActive("Word.Application").ActiveDocument.Bookmarks.Exists("general_comments") ; If bookmark exists 
 	{
		toggle++ ; Toggle to switch between current position and general comments 
		if  mod(toggle, 2) = 1
		{	
			Gosub, saveProgress
			Gosub, selectGeneralComments
		}
		if  mod(toggle, 2) = 0
			ComObjActive("Word.Application").ActiveDocument.Bookmarks("progress").Select ; Select bookmark in Word
 	}
 	Else ; If bookmark doesn't exist, insert RTF template (which contains said bookmark)
 	{
 		insertCommentsTemplate(commentsTemplate)
 		Gosub, selectGeneralComments
 	}
 	Return 


saveB1:
	ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add("b1") ; Set bookmark
	splashNotify("B1 saved") 
	Return 

saveB2:
	ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add("b2") ; Set bookmark
	splashNotify("B2 saved") 
	Return 

gotoB1:
	Try
		ComObjActive("Word.Application").ActiveDocument.Bookmarks("b1").Select ; Select bookmark in Word
	Catch
		ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add("b1") ; Set bookmark
	Return 

gotoB2:
	Try
		ComObjActive("Word.Application").ActiveDocument.Bookmarks("b2").Select ; Select bookmark in Word
	Catch
		ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add("b2") ; Set bookmark
	Return 

gotoProgress:
	Try 
		ComObjActive("Word.Application").ActiveDocument.Bookmarks("progress").Select ; Select bookmark in Word
	Catch 
		ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add("progress") ; Set bookmark
	Return 
	
gotoPaneProgress:
	Try 
		ComObjActive("Word.Application").ActiveDocument.Bookmarks("pane_progress").Select ; Select bookmark in Word
	Catch 
		ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add("pane_progress") ; Set bookmark
	Return 


gotoTableProgress:
	oWord := ComObjActive("Word.Application")
	Try 
		ComObjActive("Word.Application").ActiveDocument.Bookmarks("table_progress").Select ; Select bookmark in Word
	Catch
		ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add("table_progress") ; Set bookmark
	Return 

scrollIntoView:
	aW := ComObjActive("Word.Application").ActiveWindow
	cursor := ComObjActive("Word.Application").Selection.Range 
	firstParagraph := ComObjActive("Word.Application").ActiveDocument.Paragraphs(1).Range
	
	; Scroll to first paragrah (needed for consistency), then scroll to cursor 
	aW.ScrollIntoView(firstParagraph, True) 
	aW.ScrollIntoView(cursor, True)  
	
	; ; Scroll up a bit, so the cursor will be at the top 
	aW.SmallScroll(3,0,0,0)
 	
 	; if aW.Panes(1).View.Type = 6 ; If main window is in web view 
 	; {
	; 	if aW.View.Zoom.Percentage > 150
	; 	{
	; 		aW.SmallScroll(3,0,0,0)
	; 		splashNotify("web greater") 
	; 	}
	; 	Else 
	; 	{
	; 		aW.SmallScroll(3,0,0,0)
	; 		splashNotify("web & less") 
	; 	}
 	; }	
	; if aW.Panes(1).View.Type = 3 ; If main window is in print view 
	; {
	; 	if aW.View.DisplayPageBoundaries = 0 ; If page boundaries hidden 
	; 	{
	; 		if aW.View.Zoom.Percentage > 150
	; 		{
	; 			aW.SmallScroll(2,0,0,0)
	; 			splashNotify("print & greater") 
	; 		}
	; 		Else 
	; 		{
	; 			aW.SmallScroll(0,0,0,0)
	; 			splashNotify("print & less") 
	; 		}
	; 	}
	; 	Else
	; 	{	
	; 		aW.SmallScroll(4,0,0,0)
	; 		splashNotify("print + boundaries on") 
	; 	}
	; }
	Return 

saveTableProgress:
	oWord := ComObjActive("Word.Application") 
	oWord.ActiveDocument.Bookmarks.Add("table_progress") ; Set table progress bookmark
	Return 

saveProgress:
	ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add("progress") ; Set progress bookmark
	Return 

savePaneProgress:
	ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add("pane_progress") ; Set progress bookmark
	Return 


setBookmark:
	Loop 
	{ ; looking for a free bookmarkname
		If oWord.ActiveDocument.Bookmarks.Exists("b" A_Index) 
			continue
		temp := "b" A_Index		
		break
	}
	oWord.ActiveDocument.Bookmarks.Add( temp, oWord.Selection.Range )
	bookmark_i := oWord.ActiveDocument.Bookmarks(temp).Range.BookmarkID
	splashNotify("Set bookmark", "bottom")
Return 


deleteBookmark:
	oWord := ComObjActive("Word.Application") 
	if (bookmark_i > 0)	
	{
		thisBookmark := oWord.ActiveDocument.Bookmarks(bookmark_i).Name
		If oWord.ActiveDocument.Bookmarks.Exists(thisBookmark) 
		{
			oWord.ActiveDocument.Bookmarks(thisBookmark).Delete
			splashNotify("Bookmark deleted", "bottom")
		}
	}
	Return

jumpToBookmark: ; jump down to next bookmark
	oWord := ComObjActive("Word.Application")
	bookmarks_Count := oWord.ActiveDocument.Bookmarks.count ; count number of defined bookmarks
	If (bookmarks_Count > 0) 
	{
		If (bookmark_i < bookmarks_Count and bookmark_i > 0) 
		{
	 		bookmark_i++
			oWord.ActiveDocument.Bookmarks(Index := bookmark_i).select  ; jump to i.-bookmark
		}
		Else If (bookmark_i = bookmarks_Count) 
		{ 
			oWord.ActiveDocument.Bookmarks(Index := 1).select  ; jump to first bookmark
			bookmark_i := 1
	  	}
	  	Else 
	  	{ ; if bookmark_i = 0 or undefined 
	    	oWord.ActiveDocument.Bookmarks(Index := bookmarks_Count).select  ; jump to the last bookmark
	  		bookmark_i := bookmarks_Count 
		}
		splashNotify("Jump to bookmark", "bottom")
	}
	Return


selectGeneralComments:
	oWord := ComObjActive("Word.Application") 
	oWord.ActiveDocument.Bookmarks("general_comments").Select ; Select bookmark in Word
	oWord.Selection.StartOf(Unit:=4)  ; Select current paragraph in Word
	oWord.Selection.MoveEnd(Unit:=4)
	Send {End}
	Return

singleUppercase:
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Send, ^c 
	ClipWait
	Gosub, blueFont
	Gosub, boldFont 
	Send % ChangeCase(Clipboard,"I")
	Clipboard = %ClipSaved%  ; Restore clipboard
	Return 
insertComma:
	oWord := ComObjActive("Word.Application") 
	Gosub, boldFont 
	Gosub, blueFont
	Send, ,←
	Send, +{Left}
	oWord.Selection.Font.Subscript := 1
	Send, {Right 2}
	Return 
movePeriod:
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Send, {Right}
	Send, ←
	Send, +{Left}
	ComObjActive("Word.Application").Selection.Font.Subscript := 1
	Send, {Right}
	Send, +{Left 2}
	Gosub, boldFont 
	Gosub, blueFont
	Send, ^c 
	ClipWait
	Gosub, greyFont
	Send, {Right 2}
	; Clipboard = %ClipSaved%  ; Restore clipboard 
	Return 

insertPeriod:
	Gosub, boldFont 
	Gosub, blueFont
	Send, .{Right}
	Send, +{Right}
	Gosub, singleUppercase
	Return 

cutAndMove:
	Gosub, boldFont
	Gosub, blueFont
	Send, ^c
	Sleep 1000
	Gosub, clearFormatting
	Gosub, strikeFont
	Return 

inlineComment:
	oWord := ComObjActive("Word.Application") 
	if oWord.Selection.Type <> 1 ; If selection, format  and position cursor before inserting comment 
	{		
		; oWord.Selection.ExtendMode := False ; Turns off Word's extend mode 
		Gosub, greyFont
		Gosub, boldFont
		Gosub, wStringInfo
		if rChar = %A_Space%
		{ 
			Send {Right}
			Send {Left}
		}
		Else
			Send {Right}
	}
	Sleep 100
	Gosub, formatInlineComment
	oWord.Selection.LanguageID := 1024 ; No proofing
	oWord.Selection.Font.Name := "Courier New" 
	Send, {!} .
	Send {Left}
	oWord.Selection.LanguageID := 1024 ; No proofing
	oWord.Selection.Font.Name := "Courier New"
	Return 

wStringInfo:
	Global rChar
	Global wLen 
	Global isComma
	Global isPeriod
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	ComObjActive("Word.Application").Selection.Copy 
	StringRight, rChar, Clipboard, 1
	StringLen, wLen, Clipboard 
	if Clipboard = ,
	{
		isComma = 1
		isPeriod = 0
	}
	Else if Clipboard = .
	{
		isComma = 0
		isPeriod = 1	
	}
	Else 
	{
		isComma = 0
		isPeriod = 0
	}
	Clipboard = %ClipSaved%  ; Restore Clipboard
	Return 

formatInlineComment:
	oWord := ComObjActive("Word.Application") 
	oWord.Selection.Font.Underline := 1
	oWord.Selection.Font.StrikeThrough := 0
	oWord.Selection.Font.Bold := 0
	oWord.Selection.Font.Italic := 0
	oWord.Selection.Font.Color := -587137152 ; 16748574
	oWord.Selection.Font.Superscript := 1
	Return

underlineFont:
	ComObjActive("Word.Application").Selection.Font.Underline := 1
	Return 

courier:
	ComObjActive("Word.Application").Selection.Font.Name := "Courier New"
	Return 

spelling:
	Gosub, inlineComment
	Send, stavefejl
	Return 

selectionOrWord:
	oWord := ComObjActive("Word.Application") 
	if oWord.Selection.Type = 1 ; If no selection 
		oWord.Selection.Words(1).Select ; Select the current word
	Gosub, wStringInfo
	If rChar = %A_Space% 
		Send, +{Left}
	Return 

selectWord:
	ComObjActive("Word.Application").Selection.Words(1).Select ; Select the current word
	Return 

trailingSpace:
	trailingSpace = 0 ; Reset trailing space counter 
	Loop 
	{
		StringRight, rChar, Clipboard, 1
		If rChar = %A_Space% 
		{
			trailingSpace++
			StringTrimRight, Clipboard, Clipboard, 1
		}
		Else 
			Break 
	}
	Return 



then:
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Clipboard = →
	Sleep 100
	ComObjActive("Word.Application").Selection.PasteAndFormat(22)
	Clipboard = %ClipSaved%  ; Restore clipboard
	Return 


boldFont:
	ComObjActive("Word.Application").Selection.Font.Bold := 1
	Return 
blueFont:
	ComObjActive("Word.Application").Selection.Font.Color := 16748574  
	Return 
greyFont: 
	ComObjActive("Word.Application").Selection.Font.Color := -587137152
	Return 
redFont: 
	ComObjActive("Word.Application").Selection.Font.Color := 192
	Return 

clearFont:
	ComObjActive("Word.Application").Selection.Font.Color := -16777216
	Return 

clearFormatting:
	oWord := ComObjActive("Word.Application") 
	Gosub, clearFont 
	oWord.Selection.Font.Underline := 0
	oWord.Selection.Font.Bold := 0
	oWord.Selection.Font.StrikeThrough := 0
	oWord.Application.Selection.Range.HighlightColorIndex := 0 ; Auto (off)
	oWord.Application.Selection.Shading.BackgroundPatternColor := -16777216 ; Auto (off)
	Return

errorFont:
	Gosub, selectionOrWord
	Gosub, redFont 
	ComObjActive("Word.Application").Selection.Font.Bold := 1
	Gosub, underlineFont
	; Gosub, courier
	Send {Right}
	Return 

strikeFont:
	Gosub, selectionOrWord
 	Gosub, wStringInfo
	Gosub, greyFont ; redFont 
	Gosub boldFont 
	if (isComma || isPeriod)
	{
		Send, {Right}
		Gosub, formatInlineComment
		oWord.Selection.Font.Subscript := 1
		oWord.Selection.Font.Name := "Courier New"
		if isComma
			Send, {!} Ikke komma.
		if isPeriod 
			Send, {!} Ikke punktum.
		oWord.Selection.Font.Subscript := 0
	}
	Else 
		ComObjActive("Word.Application").Selection.Font.StrikeThrough := 1
	Return

xlSnippets:
	; Brings up a list of standard comments and inserts them 
	try 
		xl := ComObjActive("Excel.Application") 
	catch {
		xl := ComObjCreate("Excel.Application")
	}
	try
		xl.Workbooks("xlSnippets.xlsx").Activate
	catch {
		xl.Workbooks.Open(xlSnippets) ; Defined under "; === Files ==="
	}
	; Sleep 1000
	if xlSnippetsVisible = 1
		xl.Visible := True
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Gosub, xlTitles ; Display title column and return %choice% 
	if choice <> ; Run only if choice isn't empty
	{
		xl.Range("B" . choice).Select ; Select cell, instead of just copying it, so it will be easy to edit comment in Excel
		xl.Selection.Copy
		Clipboard = %Clipboard% ; Trim leading and trailing spaces. Remove formatting. 
		WinActivate, ahk_class OpusApp
	 	WinWaitActive, ahk_class OpusApp
		Send, ^v ; ComObjActive("Word.Application").Selection.PasteAndFormat(22) ; Paste without formatting  
		Send, {Right 2}
	}
	text = 
	row = 
	Clipboard = %ClipSaved%  ; Restore clipboard
	Return			


xlTitles:
	global titles = ; Clear titles 
	global choice = ; Clear choice
	Loop, % xl.ActiveSheet.UsedRange.Rows.Count  ; Loop through used columns 
		titles .= xl.Range("A" . A_Index).Value "|" ; Get column headers
	StringTrimRight, titles, titles, 1 ; Delete final |-character
	splashList_AltSubmit("", titles)
	Return


