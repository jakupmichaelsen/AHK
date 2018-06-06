#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
IniRead, browser, splashUI.ini, paths, browser
IniRead, onedrive, splashUI.ini, paths, onedrive

Global browser 

#NoTrayIcon
SendMode Input
TrayTip, %A_ScriptName%, Script loaded
AutoReload()


#F12::Gosub, gotoSublime
+F12:: ; Copy higlighted text and paste into new file in Sublime Text 
  ClipSaved = %ClipboardAll%  ; Save clipboard
  Clipboard = ; Clear clipboard 
  Send ^c 
  ClipWait
  Gosub, gotoSublime
  WinWaitActive, ahk_class PX_WINDOW_CLASS
  Send ^n 
  Sleep 100
  Send ^v 
  Clipboard = %ClipSaved%  ; Restore clipboard
  Return 

^F12::WinClose, ahk_class PX_WINDOW_CLASS
; StackEdit:--app-id=iiooodelglhkcpgbajoejffhijaclcdg
;^F6::run taskkill /F /IM fcc32.exe,,Hide
; #IfWinActive, ahk_class Chrome_WidgetWin_1
; +F11::
; 	ClipSaved = %ClipboardAll%  ; Save clipboard
; 	Clipboard = ; Clear clipboard 
; 	Send, !d 
; 	Sleep 50
; 	Send, ^c
; 	ClipWait 
; 	run  %browser% --app=%clipboard%
; 	Clipboard = %ClipSaved%  ; Restore clipboard
; 	Return 
; #IfWinActive 

#F11:: 
    IfWinExist, Messenger ahk_class Chrome_WidgetWin_1
        WinActivate
    Else
		run %browser% --profile-directory=Default --app=https://www.messenger.com/t/kristina.thastrup
	Return 
+F11::run %browser% --profile-directory=Default --app=https://inbox.google.com


; +F11:: ; Orange background in Word 
; 	oWord := ComObjActive("Word.Application") 
; 	toggle++
; 	if  mod(toggle, 2) = 1
; 	{
; 		oWord.ActiveDocument.ActiveWindow.View.DisplayBackgrounds := 1
; 		oWord.ActiveDocument.Background.Fill.ForeColor.ObjectThemeColor := 7
; 		oWord.ActiveDocument.Background.Fill.Visible := -1
; 		; oWord.ActiveDocument.Background.Fill.Solid
; 	}
; 	if  mod(toggle, 2) = 0
; 		oWord.ActiveDocument.Background.Fill.Visible := 0
; 	Return
 
#IfWinNotActive, ahk_class ENMainFrame
^F11::
    SetTitleMatchMode, RegEx 
    WinKill, ^Inbox.*jakupmichaelsen@gmail.com
    Return
#IfWinNotActive


+!F11::
	; splashNoteSmall()
	; 	Clipboard := input 
	; Return 
	splashNoteSmall()
	if input <>
	{
		Clipboard := input ; Copy input to Clipboard
		; StringLeft, fileName, input, 20 ; First 20 characters of input will be splash note filename 
		; FileAppend, %input%, %onedrive%\Markdown\Notes\sN - %fileName%.md ; Save as markdown 
	}
	Return 

!^F11::
	if Clipboard <> ; If clipboard not empty 
		defaultTxt := Clipboard
	Else
		defaultTxt = "" ; If clipboard is empty set defaulfTxt to empty string 
	splashNoteFull(20, defaultTxt)
	Clipboard := input ; Copy input to Clipboard
	Return 

; +#F10::
;     run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --incognito  --app=https://www.lectio.dk/lectio/500/forside.aspx
;     Return
; +F10:: ; openLectio()
;     run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe  --app=https://www.lectio.dk/lectio/248/SkemaNy.aspx?type=laerer&laererid=9130528370
;     Return

#F10::run lectioLinks.ahk 

^F10::
    SetTitleMatchMode, 2
    GroupAdd, Lectio, Lectio
    GroupClose, Lectio, A
    Return

#F9::run %browser% --profile-directory=Default --new-window https://workflowy.com/# 
+F9::
	splashList("WorkFlowy", "1h|1r|2r|2x")
	if choice <>
	{
		if choice = 1h
			run %browser% --profile-directory=Default --new-window https://workflowy.com/#/76ac41bb2731
		if choice = 1r
			run %browser% --profile-directory=Default --new-window https://workflowy.com/#/85b58169d668
		if choice = 2r
			run %browser% --profile-directory=Default --new-window https://workflowy.com/#/a7649304775c
		if choice = 2x
			run %browser% --profile-directory=Default --new-window https://workflowy.com/#/ae9efdf4bd9f
	}
	Return 

; openWorkFlowy()
; +#F9::run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --profile-directory=Default --new-window https://workflowy.com/s/MWMfP6dvMJ

#F8::run https://minlaering.dk/laerer/
	; openENGRAM()
+#F8::
    splashUI("r", "Direkte links til ENGRAM", "&Ordklasser|&Sætningsanalyse|&Ordstilling|&Kongruens-Verbalkongruens|&Genitiv|&Tegnsætning|&Verber|&Substantiver|&Adjektiver|&Adverbier|&Artikler|&Pronominer|&Præpositioner")
    if choice <> ; Run only if choice isn't empty
    {
        StringReplace, choice, choice, æ, ae
        StringReplace, choice, choice, &
        StringLower, choice, choice
        run, % "http://www.engram.dk/grammatik/" . choice
    }
    Return
#F7::
    run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe  --app=https://mightytext.net/web8/
    Return
+#F7:: ; DOESN'T WORK AT THE MOMENT 
    splashUI("t", "Google Timer`rHow many minutes?", 1)
    ; set timer n minutes
    url := % "https://www.google.dk/search?sourceid=chrome-psyapi2&ion=1&espv=2&ie=UTF-8&q=set%20timer%20" . input . "%20minutes"
    b = C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
    if input <>
        Run, % b --app=%url%
    Return
    ; %browser% --app=%clipboard% 
#F6::run C:\Program Files (x86)\FirstClass\fcc64.exe
^F6::
	IfWinExist, ahk_class SAWindow
		run taskkill /F /IM fcc64.exe,,Hide
	Return 
+^F6::
	WinClose, ahk_class SAWindow 
	Sleep 1000
	run C:\Program Files (x86)\FirstClass\fcc64.exe
	Return 

#F5::
    SetTitleMatchMode 2
    IfWinExist, Google Keep
            WinActivate
    Else
        run %browser% --profile-directory=Default --app=https://keep.google.com
    Return 


; #F4::
;--------------------------------------------------------
+F4:: 
	splashList("Display","Clone|Extend|Off|Only", sorted=0)  ; Pass help text and |-separated list
	if choice <>
	{
		if choice = Clone
			run DisplaySwitch.exe /clone
		if choice = Extend
			run DisplaySwitch.exe /Extend
		if choice = Off
			run DisplaySwitch.exe /internal
		if choice = Only
			run DisplaySwitch.exe /external	
	}
	Return
; ↓ Keyboard shortcuts for Display splashList 
#IfWinActive, GUI splashList, Display 
F4::
Tab::
	Send {Down}
	Return 

+F4::
+Tab::
	Send {Up}
	Return 
#IfWinActive
;--------------------------------------------------------
; #F4::openElever()
/* $F4::                                               ; Long press (> 0.5 sec) on F4 closes window
    KeyWait, F4, T0.5                               ; Wait no more than 0.5 sec for key release (also suppress auto-repeat)
    If ErrorLevel
    	{
    		; #IfWinActive ahk_class Chrome_WidgetWin_1
    		; 	Send, ^w 							; If in Chrome, close tab
    		; #IfWinActive
    			PostMessage, 0x112, 0xF060,,, A     ; Else: close window
    	}
    Else                                                ; otherwise...
        Send {F4}                                       ; ...just send F4
Return
*/
#IfWinActive, ahk_class OpusApp
!F3::
	; "Select All Text With Similar Formatting" 
	; Requires 'Select' to be in the position 1 in the Quick Access Toolbar
	Send !1s
	Return 
#IfWinActive


#F3::
	IfWinExist, Google Drive
		WinActivate
    Else
		run %browser% --profile-directory=Default --app=https://drive.google.com/drive/#recent
	Return

+F3:: ; Change case of higlighted text 
Send ^c 
ClipWait
splashText("Change case", 1, "Type is S,I,U,L, or T ") ; Prompt for type:
if input <> ;  Run only if input not empty 
{
	if input = S
		Send % ChangeCase(Clipboard,"S")
	if input = I
		Send % ChangeCase(Clipboard,"I")
	if input = U 
		Send % ChangeCase(Clipboard,"U")
	if input = L 
		Send % ChangeCase(Clipboard,"L")
	if input = T
		Send % ChangeCase(Clipboard,"T")
}
Return 

!F2::
	splashText("Lookup word", 1, "d: Dictionary.com, t: Thesaurus.com, o: ordnet.dk, `rs: Sprotin, n: Google Ngram Viewer")
	if input <>
	{
		StringReplace, input, input, æ, `%C3`%A6, All
		StringReplace, input, input, ø, `%C3`%B8, All
		StringReplace, input, input, å, `%C3`%A5, All
		if SubStr(input, 1,2) = "d "
			url = http://www.dictionary.com/browse/
		if SubStr(input, 1,2) = "t "
    		url = http://www.thesaurus.com/browse/
		if SubStr(input, 1,2) = "o "
			url = http://ordnet.dk/ddo/ordbog?query=
		if SubStr(input, 1,2) = "s "
			url = http://sprotin.fo/?p=dictionaries&_SearchDescription=1&_DictionaryPage=1&_DictionaryId=1&_SearchFor=
		if SubStr(input, 1,2) = "n "
			url = https://books.google.com/ngrams/graph?year_start=1800&year_end=2000&content=
		StringTrimLeft, input, input, 2
		StringReplace, input, input, %A_Space%, +, All
		Run %browser% --app=%url%%input%	
	}
	Return

#IfWinNotActive, ahk_class PX_WINDOW_CLASS
^F2:: ; Close dictionary windows 
	SetTitleMatchMode, 2
	GroupAdd, dictionaries, Den Danske Ordbog
	GroupAdd, dictionaries, Thesaurus.com
	GroupAdd, dictionaries, Dictionary.com
	GroupClose, dictionaries, A
	SetTitleMatchMode, 1
	Return

#F2:: ; Run scripts 
    path = %A_ScriptDir%\myScripts.txt ; Path to txt file 
    txtList(path) ; Function to parse txt file, readying if for splashUI 
    splashList("Select script to run", options)
    if choice <> ; Run only if choice isn't empty 
        Run, %A_ScriptDir%\%choice%
    Return

; +^F2:: ; Close scripts
;     path = %A_ScriptDir%\myScripts.txt ; Path to txt file 
;     txtList(path) ; Function to parse txt file, readying if for splashUI radio 
;     splashList("Close script", options)
;     if choice <> ; Run only if choice isn't empty 
;         {
;             DetectHiddenWindows, On
;             WinClose, %A_ScriptDir%\%choice% ahk_class AutoHotkey
;         }
;     TrayTip, AutoHotkey, %choice% closed
;     Return

#IfWinActive, ahk_class OpusApp
+F2::
    oWord := ComObjActive("Word.Application") 
    Try 
		oWord.Selection.Tables(1).Select
	Catch
	    oWord.Selection.StartOf(Unit:=4)  ; Select current paragraph in Word
    	oWord.Selection.MoveEnd(Unit:=4)
    Return   
#IfWinActive

; $F2::                                               ; Long press (> 0.15 sec) on F2
;     KeyWait, F2, T0.15                               ; Wait no more than 0.15 sec for key release (also suppress auto-repeat)
;     If ErrorLevel                                       ; timeout, so long press
; 		Send, #r
;     Else                                                ; otherwise...
;         Send {F2}                                       ; ...just send F2
; 	Return

F1::
    splashText("Google Search", 1, "")
    if input <> ; if input is not empty
    {
    	if SubStr(input, 1, 2) = "d "
    	{
    		StringTrimLeft, input, input, 2
    		newInput = define %input%
    		input := newInput
    	}

    	StringReplace, input, input, +, `%2B, All
    	StringReplace, input, input, %A_Space%, +, All
        Run %browser% https://www.google.com/search?q=%input%
    }
    Return

~RButton & LButton::
+F1:: ; Google highlighted text 
    ClipSaved = %ClipboardAll%  ; Save clipboard
    Clipboard = ; Clear clipboard 
    send ^c ; Copy selected text 
    ClipWait
    if SubStr(Clipboard, 1, 4) = "http"  ; If RegExMatch(Clipboard, "^(https?://|www\.)[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,3}(/\S*)?$") ; If URL
    	Run %Clipboard% ; Open URL
	Else if GetKeyState("CapsLock", "T") = 1 ; If CapsLock is on, define the word
        run http://www.google.com/search?q=define+%clipboard%
    Else ; Just google it 
        run http://www.google.com/search?q=%clipboard%
    Clipboard = %ClipSaved%  ; Restore clipboard
    Return
#F1::run %PROGRAMFILES%\AutoHotkey\WindowSpy.ahk
!F1::run myCommands.ahk


gotoSublime:
    IfWinExist, ahk_class PX_WINDOW_CLASS
        WinActivate
    Else
        run %onedrive%\Apps\Sublime Text Build 3143 x64\sublime_text.exe
    return 