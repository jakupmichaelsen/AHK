#Persistent
#NoEnv ; Significantly improves performance by not looking up empty variables as environmental variables 
#SingleInstance, force

iniRead, onedrive, splashUI.ini, paths, onedrive

Menu, Tray, Icon, %onedrive%\Code\AHK\icons\word-128.ico
Gosub, singleKeysOff
splashNotify(A_ScriptName . " loaded", "middle") 
AutoReload()

ComObjError(0)

global onedrive 
global toggle = 0

global copy2XL = 0
global screenUpdating = 0

global xlSnippetsVisible = 1
global xlPasteFormatting = 1 
global feedbackInsert = "top" ; "top" ; "bottom"

; === Files === 

global xlSnippets
global §comment
global commentsTemplate
global feedbackTemplate

xlSnippets = %onedrive%\Lektor\Grading\xlSnippets.xlsx
§comment = %onedrive%\Lektor\Grading\Templates\§ comment.rtf
commentsTemplate = %onedrive%\Lektor\Grading\Templates\comment.rtf ; 
feedbackTemplate = %onedrive%\Lektor\Grading\Templates\language.rtf ; 

if copy2XL = 1  
	TrayTip, %A_ScriptName%, NB: Copying to Excel is active, 4, 17

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
	splashNotify("minimalistFeedback", "statusBar", 600000) 
	Return

#If GetKeyState("NumLock", "T") and WinActive("ahk_class OpusApp")

; ******************
; === NAVIGATION ===
; ******************

w::
	if A_CaretY < 200
		Send, {WheelUp}
	Else 
		Send, {Up}
  Return 

a::Send, ^{Left}
s::
	if A_CaretY > 600
		Send, {WheelDown}
	Else 
		Send, {Down}
	Return 
d::Send, ^{Right}

^w::Send, ^{Up}
^a::Send, ^{Left}
^s::Send, ^{Down}
^d::Send, ^{Right}

+w::Send, +{Up}
+a::Send, +^{Left}
+s::Send, +{Down}
+d::Send, +^{Right}

+^w::Send, +^{Up}
+^a::Send, +^{Left}
+^s::Send, +^{Down}
+^d::Send, +^{Right}

; ******************
; === HOTKEYS ===
; ******************

Space::Gosub, scrollIntoView
SC02B::feedback("Genitiv") ; SC02B = ' 
-::feedback("Sammensat adjektiv")
<::feedback("False friends")
+::feedback("Præposition før that-sætning")
,::feedback("Komma")
.::feedback("Punktum")
b::feedback("Abstrakt/ generelt substantiv") ; ("Fejl ved artikler")
c::feedback("Stavning")
e::feedback("Artikler")
f::feedback("Substantivers flertal")
g::feedback("Ordvalg")
h::Gosub, blackHighlighter
i::feedback("Danisme")
j::feedback("Gradbøjning af adjektiver")
k::feedback("Kongruensfejl")
l::feedback("Ordklasse")
m::feedback("Formulering")
n::feedback("Pronominer")
o::feedback("Ordstilling")
p::feedback("Præpositioner")
q::feedbackQuery() ; Query for adding custom feedback category 
r::feedback("Proprier")
t::feedback("Uformelt")
u::feedback("Udvidet tid")
v::feedback("Verbets tider")
x::feedback("Hjælpeverber")
y::feedback("Syntaks")
z::feedback("Adj. ↔ adv.")

#If 

#IfWinActive, ahk_class OpusApp
~!^m::
	Gosub, singleKeysOff ; Toggle single-key feedback off when adding Word's own comments
	Return 

~!^e::
	Gosub, singleKeysOff ; Toggle single-key feedback off when toggling Track Changes in Word
	Return 

+^Enter::Gosub, paragraphBreak ; Gosub, removeNewLine

+^0::Gosub, clearFormatting 
F2::Gosub, toggleFeedbackAreas
+F4::Gosub, tableComment 
F6::Gosub, xlSnippets
F8::Gosub, generalComments

#If  

; ******************
; === FUNCTIONS ===
; ******************

feedback(what) {
	; Define Word objects 

	oWord := ComObjActive("Word.Application")
	wDoc := oWord.ActiveDocument

	oWord.ActiveDocument.TrackRevisions := False 

	if screenUpdating = 0
		oWord.Application.ScreenUpdating := False

	saveBookmark("progress") ; Set progress bookmark

	; All feedback tables have "table_progress" bookmark. If bookmarks doesn't exist, insert feedback table. 
 	if not wDoc.Bookmarks.Exists("feedback_table") 
 		insertFeedbackTemplate(feedbackTemplate)

	gotoBookmark("progress") ; Select progress bookmark 

	tag := what ; Copy what to tag variable 
	StringReplace, tag, tag, %A_Space%, _, All ; Replace space in tag with underscore 
	StringReplace, tag, tag, -, _, All ; Replace dash in tag with underscore 

	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 

	Gosub, selectionOrWord ; Check if something's selected, otherwise select current/closest word 
	oWord.Selection.Copy ; Copy selection 
	previewSelection := Clipboard ; Create preview of selection 
	Gosub, underlineFont ; Format selection 
	Gosub, boldFont 
	Gosub, copyWholeSentence ; Now, copy the whole sentence 
	if (what = "Komma" or what = "Punktum")
		Clipboard = %Clipboard% ; Remove formatting for "komma" and "punktum"
	gotoBookmark("progress")
	Gosub, clearFormatting ; Clear selection formatting 
	; tts(previewSelection)
	; splashNotify(what . "    " . previewSelection . "    " . SubStr(Clipboard, 1, 40), "top", 1500, 14, 200)
	splashNotify(what . "    " . previewSelection, "typing", 1500, 14, 250)
	
	if copy2XL = 1
		xlPaste(Clipboard, tag)
	
	gotoBookmark("table_progress")
	oWord.Selection.InsertRowsBelow 
	gotoBookmark("table_progress")
	
	
	oWord.Selection.TypeText(what) ; Type the category 
	oWord.Selection.MoveRight(Unit:=12) ; wdCell = 12
	oWord.Selection.PasteAndFormat(20) ; Paste sentence destination formatting 
	oWord.Selection.MoveRight(Unit:=12) ; wdCell = 12
	oWord.Selection.MoveRight(Unit:=12) ; wdCell = 12
	saveBookmark("table_progress") ; Set table progress bookmark	
	
	Clipboard = %ClipSaved%  ; Restore clipboard
	gotoBookmark("progress")
	Gosub, scrollIntoView
 	Send, {CtrlUp}{Right}
	Gosub, saveDoc 
	oWord.Application.ScreenUpdating := True
	Gosub, singleKeysOn
}

feedbackQuery() {
	global input 
	splashInput("Error type?")
	if input <> ; If input not empty 
		feedback(input)
}

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

saveBookmark(bookmark){
	ComObjActive("Word.Application").ActiveDocument.Bookmarks.Add(bookmark) ; Update terminology bookmark
}

gotoBookmark(bookmark){
	ComObjActive("Word.Application").ActiveDocument.Bookmarks(bookmark).Select ; Select bookmark in Word
}


insertFeedbackTemplate(name){
	wSelection := ComObjActive("Word.Application").Selection
 	toggle = 1 ; Reset toggle counter for toggling between bookmarks 
 	if feedbackInsert = bottom
 	{
 		Gosub, gotoEnd
		wSelection.InsertBreak(2) ;  wdSectionBreakNextPage 
 	}
 	else if feedbackInsert = top
 	{
 		Gosub, gotoHome
		wSelection.InsertBreak(2) ;  wdSectionBreakNextPage 
		Gosub, gotoHome
		; oWord.Selection.PageSetup.Orientation := 1 ;  wdOrientLandscape = 1
 	}
 	else 
 	{	
		Msgbox Set %feedbackInsert% to either "top" or "bottom". Script will now reload!
 		Reload
 		Sleep 100
 	}
	wSelection.LanguageID := 1024 ; No proofing
	wSelection.PageSetup.TopMargin := 15
	wSelection.PageSetup.BottomMargin := 15
	wSelection.InsertFile(feedbackTemplate)
}

; ******************
; === GOSUBS ===
; ******************

blackHighlighter:
	oWord.ActiveDocument.TrackRevisions := False
	Gosub, selectionOrWord
	oWord := ComObjActive("Word.Application")
	oWord.Application.Selection.Range.HighlightColorIndex := 1 ; Black 
	oWord.Application.Selection.Range.Font.Color := 0xFFFFFF ; White 
	oWord.Selection.Font.Name := "Courier New" 
	Send {Right}
	Gosub, saveDoc
	Return 

saveDoc:
	aDoc := ComObjActive("Word.Application").ActiveDocument
	aDoc.Save
	aDoc.EmbedTrueTypeFonts := True 
	Return 

toggleFeedbackAreas:
	toggle := Mod( toggle, 3 )+1
	if  toggle = 1
	{
		Gosub, singleKeysOn
		gotoBookmark("progress")
		Gosub, scrollIntoView
	}
	if toggle = 2 
	{
		if ComObjActive("Word.Application").ActiveDocument.Bookmarks.Exists("table_progress") 
		{
			Gosub, singleKeysOff
			gotoBookmark("table_progress")
			Gosub, scrollIntoView
			; msgbox % "A_CaretY: " . A_CaretY . " A_CaretX: " . A_CaretX
		}
		Else 
		{	
			Gosub, singleKeysOn
			gotoBookmark("progress")
		}
	}
	if toggle = 3
	{
		if ComObjActive("Word.Application").ActiveDocument.Bookmarks.Exists("general_comments") 
		{
			Gosub, singleKeysOff
			gotoBookmark("general_comments")
		}
		Else
		{
			Gosub, singleKeysOn
			Gosub, gotoHome
		} 
	}
	Return 

selectSentence:
	oWord := ComObjActive("Word.Application")
	oWord.Selection.MoveLeft(Unit:=3)  ; Move to start of sentence
	oWord.Selection.ExtendMode := True ; Turns on extend mode 
	oWord.Selection.MoveRight(Unit:=3) ; Expand selection to whole sentence
	oWord.Selection.ExtendMode := False ; Turns off extend mode 
	Return 

tableComment: ; Table paragraph comment - cuts selection and pastes it into table from RTF file 
	oWord := ComObjActive("Word.Application") 
 	oWord.ActiveDocument.TrackRevisions := False
 	ClipSaved = %ClipboardAll%  ; Save clipboard
 	Clipboard = ; Clear clipboard 
	Send ^x 
	; Insert template 
	§comment = %onedrive%\Lektor\Grading\Templates\§ comment.rtf
	oWord.Selection.InsertFile(§comment)
	; Paste selection  
    oWord.Application.gotoBookmark("paste_here").Select ; Select bookmark in Word
	oWord.Selection.PasteAndFormat(16) ; Paste selection, maintaining formatting
	oWord.Application.gotoBookmark("paste_here").Delete
	Send, {BackSpace}
    oWord.Application.gotoBookmark("paragraph_comment").Select ; Select bookmark in Word
    oWord.Application.gotoBookmark("paragraph_comment").Delete
    Clipboard = %ClipSaved%  ; Restore clipboard
    Return

paragraphBreak:
	oWord := ComObjActive("Word.Application")
	oWord.ActiveDocument.TrackRevisions := True
	Send, {Enter}
	oWord.Selection.TypeText("⎶")
	oWord.Selection.LanguageID := 1033 ; wdEnglishUS
	oWord.Selection.TypeText("Paragraph break")
	Send, {Esc}
	oWord.ActiveDocument.TrackRevisions := False 
	Return 

generalComments: 
	Gosub, singleKeysOff
 	; General comment at the top of the document
 	wDoc := ComObjActive("Word.Application").ActiveDocument
 	wSelection := ComObjActive("Word.Application").Selection

	wDoc.TrackRevisions := False 

 	if wDoc.Bookmarks.Exists("general_comments") ; If bookmark exists 
 	{
		toggle++ ; Toggle to switch between current position and general comments 
		if  mod(toggle, 2) = 1
		{		
			saveBookmark("progress")
			gotoBookmark("general_comments")
		}
		if  mod(toggle, 2) = 0
			saveBookmark("general_comments")
			gotoBookmark("progress")
 	}
 	Else 
 	{
 		Gosub, gotoHome
		Send, {Enter}
 		Gosub, gotoHome
 		; wDoc.Range.ParagraphFormat.SpaceAfter := 12
 		wSelection.ParagraphFormat.Alignment := 0 ; wdAlignParagraphLeft = 0
		wSelection.Font.Name := "Rockwell Extra Bold"
		wSelection.Font.Size := 14
		wDoc.Range.ParagraphFormat.SpaceAfter := 12
 		wSelection.TypeText("”")
 		gosub, dropCap 
 		; wSelection.ParagraphFormat.LineSpacing := 12
		wSelection.LanguageID := 1033 ; English (United States)
		saveBookmark("general_comments")
 	}
 	Return 

dropCap: 
	dropCap := ComObjActive("Word.Application").Selection.Paragraphs(1).DropCap
	dropCap.Position := 1  ; wdDropNormal = 1
	dropCap.LinestoDrop := 2 
	dropCap.DistanceFromText := 2
	dropCap.FontName := "Rockwell Extra Bold"
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
	ComObjActive("Word.Application").ActiveDocument.Bookmarks("progress") ; Select progress bookmark 
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

gotoHome:
	ComObjActive("Word.Application").Selection.HomeKey(Unit := 6) ; wdStory = 6
	Return
gotoEnd:
    ComObjActive("Word.Application").Selection.EndKey(Unit := 6) ; wdStory = 6
    Return 

    copyWholeSentence:
	oWord := ComObjActive("Word.Application") 
	saveBookmark("progress") ; Set progress bookmark
	oWord.Selection.MoveLeft(Unit:=3)  ; Move to start of sentence
	oWord.Selection.ExtendMode := True ; Turns on extend mode
	oWord.Selection.MoveRight(Unit:=3) ; Expand selection to whole sentence
	oWord.Selection.ExtendMode := False ; Turns off extend mode
	oWord.Selection.Copy 
	; Sleep 100
	Return 


scrollIntoView:
	aW := ComObjActive("Word.Application").ActiveWindow
	cursor := ComObjActive("Word.Application").Selection.Range 
	firstParagraph := ComObjActive("Word.Application").ActiveDocument.Paragraphs(1).Range
	
	; Scroll to first paragrah (needed for consistency), then scroll to cursor 
	aW.ScrollIntoView(firstParagraph, True) 
	aW.ScrollIntoView(cursor, True)  
	
	; ; Scroll up a bit, so the cursor will be at the top 
	; aW.SmallScroll(6,0,0,0)

	Send, {Up}{Home}
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

underlineFont:
	ComObjActive("Word.Application").Selection.Font.Underline := 1
	Return 

courier:
	ComObjActive("Word.Application").Selection.Font.Name := "Courier New"
	Return 

selectionOrWord:
	oWord := ComObjActive("Word.Application") 
	if oWord.Selection.Type = 1 ; If no selection 
		oWord.Selection.Words(1).Select ; Select the current word
	Gosub, wStringInfo
	If rChar = %A_Space% 
	{
		Send, +{Left}
		Sleep 100
	}
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
	Return 

clearFormatting:
	wSelection := ComObjActive("Word.Application").Selection
	wSelection.Font.Color := -16777216
	wSelection.Font.Underline := 0
	wSelection.Font.Bold := 0
	wSelection.Font.StrikeThrough := 0
	wSelection.Range.HighlightColorIndex := 0 ; Auto (off)
	wSelection.Shading.BackgroundPatternColor := -16777216 ; Auto (off)
	wSelection.Range.Font.Color := 1
	Return

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

