; ============= AUTOEXEC SECTION  =============
#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
SendMode Input
AutoReload()
Menu, Tray, Icon, D:\Dropbox\Code\AHK\icons\word-128.ico
pX := A_ScreenWidth-330
pY := A_ScreenHeight-70

; run tabEditComments.ahk ; Script for navigatingn comment bubbles with TAB 
run wordColorFeedback.ahk ; Script for highlighting using the number keys 

try {
    oWord := ComObjActive("Word.Application")
}
catch {
    MsgBox,16, %A_ScriptName%, No active MS Word document.`r`rReloading....
    Reload
}


; ===============================================
global oWord := ComObjActive("Word.Application") ; Word object 
global toggle := 0 ; Global variable used for toggling 


; oWord.Options.CheckSpellingAsYouType := False
; oWord.Options.CheckGrammarAsYouType := False
; ===============================================

; ============= TOOLTIP SECTION =============
; Display tooltip with keyboard shortcuts:
; SysGet, res, Monitor ; Get screen resolution 
; ToolTip, 
; (
; %A_ScriptName% (re)loaded.
; 
;     F2: Prompt for heading and text before inserting comment. Uses highlighted text as default heading. 
;     F3: Brings up a list of standard comments from an Excel sheet and inserts them
;     F4: Prompts for comment and inserts it using current author as heading
;     
;     CapsLock: Toggle revisions tracking 
;     Ctrl + Space: Clear formatting
; ), resRight, 0
; SetTimer, RemoveToolTip, 5000
; return
; RemoveToolTip:
; SetTimer, RemoveToolTip, Off
; ToolTip
; return

; ============= WORD SECTION =============
#IfWinActive, ahk_class OpusApp ; if Word is active 
"::
	if oWord.Selection.Type = 1
    	Send ”
    Else
    {

    	ClipSaved = %ClipboardAll%  ; Save clipboard
    	Clipboard = ; Clear clipboard 
    	quote = 
		oWord.Application.Selection.Copy 
		Clipboard = %Clipboard%
    	quote .= "”" . Clipboard
    	quote .= "”"
    	Send %quote%
        Clipboard = %ClipSaved%  ; Restore clipboard
    } 
    Return
     
+F3::
	; Edit header 
 	oWord.ActiveDocument.TrackRevisions := False
 	SetScrollLockState, Off
	oWord.ActiveWindow.ActivePane.View.SeekView := 9
	Send ^{End}
	Return 

+^Space::oWord.Selection.ClearFormatting


CapsLock::oWord.ActiveDocument.TrackRevisions := !oWord.ActiveDocument.TrackRevisions

F2:: 
	; Prompt for heading and text before inserting comment. Uses highlighted text as default heading
 	oWord.ActiveDocument.TrackRevisions := False
 	SetScrollLockState, Off
    author := oWord.Application.UserName  ; Get current author
    ClipSaved = %ClipboardAll%  ; Save clipboard
    Clipboard = ; Clear clipboard 
    Try
        oWord.Application.Selection.Copy ; Copy the current word 
    Catch {
        oWord.Application.Selection.Words(1).Select ; Select the current word 
        oWord.Application.Selection.Copy ; Copy the current word 
    }
    splashText("Enter heading", 1, , Clipboard)   ; Prompt for heading (clipboard as default text)
    if input <> ; If input not empty
    {
        Clipboard = %ClipSaved%  ; Restore clipboard
        oWord.Application.UserName := input ; input becomes new heading 
        splashText("Your comment", 10, "'x' = cancel")   ; Prompt for comment 
        if input <> x                 ; If input not "x" (type x to cancel the comment) 
            wComment(input)           ; Insert comment  
        oWord.Application.UserName := author ; Restore original author 
    }
    Return

Browser_Forward::
F3:: 
	try 
    	xl := ComObjActive("Excel.Application") 
	catch {
    	MsgBox,16, %A_ScriptName%, No active Excel workbook found.`r`r
		Return 
	}
 	oWord.ActiveDocument.TrackRevisions := False
 	SetScrollLockState, Off
	refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z") ; array of column letters
	; Brings up a list of standard comments and inserts them 
    Top:
    choice =
    categorySheets(xl) ; Show selection of worksheets and choose which to activate 
    if choice <> ; Run only if choice isn't empty
    {
        xl.Worksheets(choice).Activate
        titleColumn(xl) ; Display title column and return %choice% 
        if choice <> ; Run only if choice isn't empty
        {
        	choice++ ; To skip header row 
            Loop, % xl.ActiveSheet.UsedRange.Columns.Count 
            { ; Loop through used columns 
                xl.Range(refColumns[A_Index] . choice).Select
                cell := xl.Selection.Value
                
                if A_Index > 1 ; Skip header-row
                {
                    if cell <>
                    {
                        if A_Index = 2
                        {
                            ; FileRead, Clipboard, *c %A_ScriptDir%\clipboard.tmp
                            xl.Selection.Copy
                            text := ClipboardAll
                        }
                        if A_Index = 3
                            header := cell
                        if A_Index = 4
                            URL := cell
                        if A_Index = 5
                            title := cell
                    }
                }
            }
            xlWordComment(text, header, URL, title) ; Add comment 
        }
        Else
            Gosub, Top

        header = 
        text = 
        URL =
        title = 
        row = 
    }
	Return            

F4:: 
	; Prompt for heading and text before inserting comment. Uses highlighted text as default heading
 	oWord.ActiveDocument.TrackRevisions := False
 	SetScrollLockState, Off
    author := oWord.Application.UserName  ; Get current author
    ClipSaved = %ClipboardAll%  ; Save clipboard
    Clipboard = ; Clear clipboard 
    Try
        oWord.Application.Selection.Copy ; Copy selected text  
    Catch {
        oWord.Application.Selection.Words(1).Select ; Or select the current word 
        oWord.Application.Selection.Copy ; Then copy it 
    }
	oWord.Application.UserName := Clipboard ; Clipboard becomes new heading 
	splashText("Your comment", 10, "'x' = cancel")   ; Prompt for comment 
	if input <> x                 ; If input not "x" (type x to cancel the comment) 
	    wComment(input)           ; Insert comment  
	oWord.Application.UserName := author ; Restore original author 
	Clipboard = %ClipSaved%  ; Restore clipboard
    Return

F5:: 
 	oWord.ActiveDocument.TrackRevisions := False
 	SetScrollLockState, Off
	; Prompts for comment and inserts it using current author 
    splashText("Comment", 10, "Comment using current author")   ; Prompt for comment 
    if input <> 
        wComment(input)           ; Insert comment  
    Return 

F6::
	; Table paragraph comment  
	; Cuts selection and pastes it into table from RTF file 

 	oWord.ActiveDocument.TrackRevisions := False
 	SetScrollLockState, Off
	Send ^x 
	
	; Insert template 
	dir = D:\Dropbox\Adjunkt\Grading\Templates 
	file = %dir%\§ comment.rtf
	oWord.Selection.InsertFile(FileName:=file)


	; Paste selection  
    oWord.Application.ActiveDocument.Bookmarks("paste_here").Select ; Select bookmark in Word
	oWord.Selection.PasteAndFormat(16) ; Paste selection, maintaining formatting
	oWord.Application.ActiveDocument.Bookmarks("paste_here").Delete
	Send, {BackSpace}
	Return

F8:: 
 	; General comment at the top of the document 
 	oWord.ActiveDocument.TrackRevisions := False
 	SetScrollLockState, Off
 	if oWord.Application.ActiveDocument.Bookmarks.Exists("general_comments") ; If bookmark exists 
 	{
		toggle++ ; Toggle to switch between current position and general comments 
		if  mod(toggle, 2) = 1
 			Gosub, general_comments
		if  mod(toggle, 2) = 0
			oWord.Application.ActiveDocument.Bookmarks("progress").Select ; Select bookmark in Word
 	}
 	Else ; If bookmark doesn't exist, insert RTF template (which contains said bookmark)
 	{
 		Gosub, embed_fonts ; Set Word settins for embedding fonts 
 		toggle = 1 ; Reset toggle counter for toggling between bookmarks in the above condition 
 		Send ^{Home}
		dir = D:\Dropbox\Adjunkt\Grading\Templates
		file = %dir%\Comment + Grade.rtf
		oWord.Selection.InsertFile(FileName:=file)
		Gosub, general_comments
 	} 
 	Return 

suspend_ColorFeedback:
	DetectHiddenWindows On  ; Allows a script's hidden main window to be detected.
	PostMessage, 0x111, 65403,,, %A_ScriptDir%\wordColorFeedback.ahk ; Pause script 
	Return 

general_comments:
	oWord.Application.ActiveDocument.Bookmarks.Add("progress")
	oWord.Application.ActiveDocument.Bookmarks("general_comments").Select ; Select bookmark in Word
	oWord.Selection.StartOf(Unit:=4)  ; Select current paragraph in Word
	oWord.Selection.MoveEnd(Unit:=4)
    Send {End} 
	Return

embed_fonts:
	oWord.ActiveDocument.EmbedTrueTypeFonts := True
	oWord.ActiveDocument.SaveSubsetFonts := True
	oWord.ActiveDocument.DoNotEmbedSystemFonts := True
	Return 
