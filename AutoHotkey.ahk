#NoEnv ; Significantly improves performance by not looking up empty variables as environmental variables 
#InstallMouseHook
#InstallKeybdHook
#SingleInstance, force
#NoTrayIcon

splashNotify(A_ScriptName " reloaded") 
IniRead, browser, splashUI.ini, paths, browser
IniRead, onedrive, splashUI.ini, paths, onedrive
IniRead, user, splashUI.ini, paths, user
IniRead, ahk, splashUI.ini, paths, ahk

AutoReload()
SendMode, Input
SetTitleMatchMode, 2 

run %A_ScriptDir%\WinTrigger.ahk
run %ahk%\TextExpansions.ahk
run %A_ScriptDir%\functionKeys.ahk
run %A_ScriptDir%\centerMouse.ahk
run %ahk%\WindowPadX\WindowPadX.ahk
run %A_ScriptDir%\emoji.ahk
run %A_Scriptdir%\wSuggestions.ahk
run %A_ScriptDir%\mdScreenshots.ahk
run %A_ScriptDir%\openZips.ahk



/*
																																 /$$
																																|__/
	/$$$$$$   /$$$$$$  /$$$$$$/$$$$   /$$$$$$   /$$$$$$   /$$$$$$  /$$ /$$$$$$$   /$$$$$$   /$$$$$$$
 /$$__  $$ /$$__  $$| $$_  $$_  $$ |____  $$ /$$__  $$ /$$__  $$| $$| $$__  $$ /$$__  $$ /$$_____/
| $$  \__/| $$$$$$$$| $$ \ $$ \ $$  /$$$$$$$| $$  \ $$| $$  \ $$| $$| $$  \ $$| $$  \ $$|  $$$$$$
| $$      | $$_____/| $$ | $$ | $$ /$$__  $$| $$  | $$| $$  | $$| $$| $$  | $$| $$  | $$ \____  $$
| $$      |  $$$$$$$| $$ | $$ | $$|  $$$$$$$| $$$$$$$/| $$$$$$$/| $$| $$  | $$|  $$$$$$$ /$$$$$$$/
|__/       \_______/|__/ |__/ |__/ \_______/| $$____/ | $$____/ |__/|__/  |__/ \____  $$|_______/
																						| $$      | $$                     /$$  \ $$
																						| $$      | $$                    |  $$$$$$/
																						|__/      |__/                     \______/
=== remappings === 
*/
; #IfWinNotActive, ahk_class PX_WINDOW_CLASS 
; ::"::”
; #IfWinNotActive
; ^ð::^

; SC05B::Send, {BackSpace} ; ; Disable WIFI FN-key toggle (+ for some reason this FN-key sends a chracter) ; 
;;;; Doesn't work because it also disables "external monitor"
; 
; SC056::\
; <+SC056::Send, <
; >+SC056::Send, >

+#Up::Volume_Up
+#Down::Volume_Down



#IfWinActive, ahk_class AutoHotkeyGUI, AHK Rubric.ahk ; Remap PgUp/Dn keys in AHK Rubric
PgDn::PgUp
PgUp::PgDn
#IfWinActive

#IfWinActive, ahk_class MSPaintApp ; Remap zoom keyboard shortcuts in MS Paint
^SC00C::Send {Ctrl down}{PgUp down} ; zoom in with ctrl +
^SC035::Send {Ctrl down}{PgDn down} ; zoom out with ctrl -
#IfWinActive

	

/*
 /$$      /$$                     /$$       /$$$$$$$$ /$$
| $$  /$ | $$                    | $$      | $$_____/| $$
| $$ /$$$| $$  /$$$$$$   /$$$$$$ | $$   /$$| $$      | $$  /$$$$$$  /$$  /$$  /$$ /$$   /$$
| $$/$$ $$ $$ /$$__  $$ /$$__  $$| $$  /$$/| $$$$$   | $$ /$$__  $$| $$ | $$ | $$| $$  | $$
| $$$$_  $$$$| $$  \ $$| $$  \__/| $$$$$$/ | $$__/   | $$| $$  \ $$| $$ | $$ | $$| $$  | $$
| $$$/ \  $$$| $$  | $$| $$      | $$_  $$ | $$      | $$| $$  | $$| $$ | $$ | $$| $$  | $$
| $$/   \  $$|  $$$$$$/| $$      | $$ \  $$| $$      | $$|  $$$$$$/|  $$$$$/$$$$/|  $$$$$$$
|__/     \__/ \______/ |__/      |__/  \__/|__/      |__/ \______/  \_____/\___/  \____  $$
																																									/$$  | $$
																																								 |  $$$$$$/
																																									\______/
WorkFlowy
*/
; 
; #IfWinActive, WorkFlowy ahk_class Chrome_WidgetWin_1
; PgDn::+!0
; PgUp::+!9
; #IfWinActive

/*
	/$$$$$$                                /$$                   /$$           /$$$$$$$                           /$$$$$$$   /$$$$$$
 /$$__  $$                              | $$                  | $$          | $$__  $$                         | $$__  $$ /$$__  $$
| $$  \ $$  /$$$$$$$  /$$$$$$   /$$$$$$ | $$$$$$$   /$$$$$$  /$$$$$$        | $$  \ $$ /$$$$$$   /$$$$$$       | $$  \ $$| $$  \__/
| $$$$$$$$ /$$_____/ /$$__  $$ /$$__  $$| $$__  $$ |____  $$|_  $$_/        | $$$$$$$//$$__  $$ /$$__  $$      | $$  | $$| $$
| $$__  $$| $$      | $$  \__/| $$  \ $$| $$  \ $$  /$$$$$$$  | $$          | $$____/| $$  \__/| $$  \ $$      | $$  | $$| $$
| $$  | $$| $$      | $$      | $$  | $$| $$  | $$ /$$__  $$  | $$ /$$      | $$     | $$      | $$  | $$      | $$  | $$| $$    $$
| $$  | $$|  $$$$$$$| $$      |  $$$$$$/| $$$$$$$/|  $$$$$$$  |  $$$$/      | $$     | $$      |  $$$$$$/      | $$$$$$$/|  $$$$$$/
|__/  |__/ \_______/|__/       \______/ |_______/  \_______/   \___/        |__/     |__/       \______/       |_______/  \______/


*/

#IfWinActive, ahk_class AcrobatSDIWindow
; !F4::send, +^h

^SC00C::^0
^SC035::^1
WheelLeft::Left
WheelRight::Right 

#IfWinActive


/*
																				 /$$$$$$$$ /$$ /$$
																				| $$_____/|__/| $$
	/$$$$$$   /$$$$$$   /$$$$$$  /$$$$$$$ | $$       /$$| $$  /$$$$$$       /$$$$$$  /$$   /$$  /$$$$$$
 /$$__  $$ /$$__  $$ /$$__  $$| $$__  $$| $$$$$   | $$| $$ /$$__  $$     /$$__  $$|  $$ /$$/ /$$__  $$
| $$  \ $$| $$  \ $$| $$$$$$$$| $$  \ $$| $$__/   | $$| $$| $$$$$$$$    | $$$$$$$$ \  $$$$/ | $$$$$$$$
| $$  | $$| $$  | $$| $$_____/| $$  | $$| $$      | $$| $$| $$_____/    | $$_____/  >$$  $$ | $$_____/
|  $$$$$$/| $$$$$$$/|  $$$$$$$| $$  | $$| $$      | $$| $$|  $$$$$$$ /$$|  $$$$$$$ /$$/\  $$|  $$$$$$$
 \______/ | $$____/  \_______/|__/  |__/|__/      |__/|__/ \_______/|__/ \_______/|__/  \__/ \_______/
					| $$
					| $$
					|__/

openFile.exe
*/

#o::run %A_ScriptDir%\openFile.ahk ; run Windows open file dialog window
; #IfWinActive, Select File - openFile.ahk ahk_class #32770
; Tab::Down
; <::\
; #IfWinActive

/*

	/$$$$$$   /$$$$$$  /$$$$$$$$ /$$$$$$$$        /$$$$$$                                 /$$     /$$                               /$$ /$$   /$$
 /$$__  $$ /$$__  $$|__  $$__/|__  $$__/       /$$__  $$                               | $$    |__/                              | $$|__/  | $$
| $$  \__/| $$  \__/   | $$      | $$         | $$  \__//$$   /$$ /$$$$$$$   /$$$$$$$ /$$$$$$   /$$  /$$$$$$  /$$$$$$$   /$$$$$$ | $$ /$$ /$$$$$$   /$$   /$$
|  $$$$$$ |  $$$$$$    | $$      | $$         | $$$$   | $$  | $$| $$__  $$ /$$_____/|_  $$_/  | $$ /$$__  $$| $$__  $$ |____  $$| $$| $$|_  $$_/  | $$  | $$
 \____  $$ \____  $$   | $$      | $$         | $$_/   | $$  | $$| $$  \ $$| $$        | $$    | $$| $$  \ $$| $$  \ $$  /$$$$$$$| $$| $$  | $$    | $$  | $$
 /$$  \ $$ /$$  \ $$   | $$      | $$         | $$     | $$  | $$| $$  | $$| $$        | $$ /$$| $$| $$  | $$| $$  | $$ /$$__  $$| $$| $$  | $$ /$$| $$  | $$
|  $$$$$$/|  $$$$$$/   | $$      | $$         | $$     |  $$$$$$/| $$  | $$|  $$$$$$$  |  $$$$/| $$|  $$$$$$/| $$  | $$|  $$$$$$$| $$| $$  |  $$$$/|  $$$$$$$
 \______/  \______/    |__/      |__/         |__/      \______/ |__/  |__/ \_______/   \___/  |__/ \______/ |__/  |__/ \_______/|__/|__/   \___/   \____  $$
																																																																										/$$  | $$
																																																																									 |  $$$$$$/
																																																																										\______/
*/
+^Delete::
	Send, +{End}{Delete}
	Return


#If !WinActive("ahk_class PX_WINDOW_CLASS") AND !WinActive("ahk_exe PDFXEdit.exe")
+^d::
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Sleep, 100 
	Send, {Home}
	Sleep, 100 
	Send, +{End}
	Sleep, 100 
	Send, ^c
	Send, {End}
	Sleep, 100 
	Send, {Enter}
	Send, ^v
	Sleep, 100
	Clipboard = %ClipSaved%  ; Restore clipboard
	Return


+^k::
	Send, {Home}
	Send, +{End}{Delete}
	; Send, {Delete}
	Return
#IfWinNotActive

#IfWinNotActive, WorkFlowy ahk_class Chrome_WidgetWin_1
+^BackSpace:: ; Delete whole line
	Send, +{Home}{Delete}
	Return
#IfWinNotActive

/*
 /$$$$$$$$ /$$                       /$$      /$$$$$$  /$$
| $$_____/|__/                      | $$     /$$__  $$| $$
| $$       /$$  /$$$$$$   /$$$$$$$ /$$$$$$  | $$  \__/| $$  /$$$$$$   /$$$$$$$ /$$$$$$$
| $$$$$   | $$ /$$__  $$ /$$_____/|_  $$_/  | $$      | $$ |____  $$ /$$_____//$$_____/
| $$__/   | $$| $$  \__/|  $$$$$$   | $$    | $$      | $$  /$$$$$$$|  $$$$$$|  $$$$$$
| $$      | $$| $$       \____  $$  | $$ /$$| $$    $$| $$ /$$__  $$ \____  $$\____  $$
| $$      | $$| $$       /$$$$$$$/  |  $$$$/|  $$$$$$/| $$|  $$$$$$$ /$$$$$$$//$$$$$$$/
|__/      |__/|__/      |_______/    \___/   \______/ |__/ \_______/|_______/|_______/

FirstClass
*/

#IfWinActive, ahk_class SAWindow
;!F4::send, +^h
/*
Left::MsgBox,
Backspace::
	Return
;  Shift & BackSpace::
;  	Send {BackSpace}
;  	Return
Alt & BackSpace::
	Send, !svo
	Return
Alt & Left::
	MouseClickDrag, left, 250, 40, 250, 250
	MouseClick, left, 539, 405
	; MouseMove, 539, 420
	; MouseClick, left, 539, 430
	Return

Alt & Right::
	MouseClickDrag, left, 250, 40, 250, 250
	MouseClick, left, 539, 420
	Return
*/
#IfWinActive

/*
					 /$$ /$$           /$$                                           /$$
					| $$|__/          | $$                                          | $$
	/$$$$$$$| $$ /$$  /$$$$$$ | $$$$$$$   /$$$$$$   /$$$$$$   /$$$$$$   /$$$$$$$
 /$$_____/| $$| $$ /$$__  $$| $$__  $$ /$$__  $$ |____  $$ /$$__  $$ /$$__  $$
| $$      | $$| $$| $$  \ $$| $$  \ $$| $$  \ $$  /$$$$$$$| $$  \__/| $$  | $$
| $$      | $$| $$| $$  | $$| $$  | $$| $$  | $$ /$$__  $$| $$      | $$  | $$
|  $$$$$$$| $$| $$| $$$$$$$/| $$$$$$$/|  $$$$$$/|  $$$$$$$| $$      |  $$$$$$$
 \_______/|__/|__/| $$____/ |_______/  \______/  \_______/|__/       \_______/
									| $$
									| $$
									|__/
*/


^!i:: ; Load image/clips to clipboard
	clipTop:
	splashList("Insert data", "rtf|image|clip|unicode", unsorted)
	if choice <> ; Run only if choice isn't empty
		{
				ClipSaved = %ClipboardAll%  ; Save clipboard
			Clipboard = ; Clear clipboard 
				if choice = image
				{
						dir = %A_ScriptDir%\icons 
						splashDir(dir)
						splashList("Choose image icon to load", items)
						if choice <> ; Run only if choice isn't empty
						{
						img = %dir%\%choice%
						CopyImage(img)
						msgbox %img%
								Send ^v
					}
				Gosub, clipTop
				}
				if choice = clip
				{
					dir = %A_ScriptDir%\Clips
					splashDir(dir)
					splashList("Choose saved clip to load into clipboard", items)
					if choice <>
						{
						FileRead, Clipboard, *c %A_ScriptDir%\Clips\%choice% ; Note the use of *c, which must precede the filename
								Send ^v
						}
			Else
				Gosub, clipTop 
				}
				if choice = rtf 
				{
						dir = %onedrive%\Adjunkt\Grading\Templates
						splashDir(dir, "*.rtf")
						splashList("Select RTF-file to insert", items)
						if choice <> ; Run only if choice isn't empty
						{
								oWord := ComObjActive("Word.Application")
								oWord.ActiveDocument.TrackRevisions := False
								file = %dir%\%choice%
								oWord.Selection.InsertFile(FileName:=file)
						}
			Else
				Gosub, clipTop
				}
				if choice = unicode 
				{	
					dir = %A_ScriptDir%\Clips\Unicode 
						splashDir(dir)
						splashList(" ", items)
						if choice <> ; Run only if choice isn't empty
						{
							StringRight, Clipboard, choice, 1
				; FileRead, Clipboard, *c %A_ScriptDir%\Clips\Unicode\%choice% ; Note the use of *c, which must precede the filename
					if WinActive("ahk_class OpusApp")
						ComObjActive("Word.Application").Selection.PasteAndFormat(22)
					Else
						Send ^v
						}
			Else
				Gosub, clipTop
				}
				Clipboard = %ClipSaved%  ; Restore clipboard
		}
		Return

+!i:: ; Insert Unicode
	dir = %A_ScriptDir%\Clips\Unicode 
	splashDir(dir)
	splashList("", items)
	if choice <> ; Run only if choice isn't empty
	{
			ClipSaved = %ClipboardAll%  ; Save clipboard
			Clipboard = ; Clear clipboard 

		FileRead, Clipboard, *c %A_ScriptDir%\Clips\Unicode\%choice% ; Note the use of *c, which must precede the filename
			Send ^v
			Sleep 100
			Clipboard = %ClipSaved%  ; Restore clipboard
	}
	Return


^!h:: ; Add hotstring to files 
		splashList("Save strings or clips", "clip|Unicode|TextExpansions.ahk|GradingPapers.ahk|Kristina.txt|AutoHotkey.ahk|\Kollektor\Kommentar.txt")
		if choice <> ; Run only if input isn't empty
		{
				if choice = clip
				{
						splashText("Name of clip", 1)
						if input <>
								FileAppend, %ClipboardAll%, %A_ScriptDir%\Clips\%input% ; The file extension does not matter.
				}
				Else
				if choice = Unicode
				{
						splashText("Name of unicode", 1)
						if input <>
								FileAppend, %ClipboardAll%, %A_ScriptDir%\Clips\Unicode\%input% ; The file extension does not matter.
				}
				Else
				{ 
						splashText("Insert hotstring", 10)
						if input <> ; Run only if input isn't empty
							FileAppend, %input%`r, %A_ScriptDir%\%choice%
				}
		}
		Return

/*
	/$$$$$$  /$$$$$$$  /$$        /$$$$$$   /$$$$$$  /$$   /$$
 /$$__  $$| $$__  $$| $$       /$$__  $$ /$$__  $$| $$  | $$
| $$  \__/| $$  \ $$| $$      | $$  \ $$| $$  \__/| $$  | $$
|  $$$$$$ | $$$$$$$/| $$      | $$$$$$$$|  $$$$$$ | $$$$$$$$
 \____  $$| $$____/ | $$      | $$__  $$ \____  $$| $$__  $$
 /$$  \ $$| $$      | $$      | $$  | $$ /$$  \ $$| $$  | $$
|  $$$$$$/| $$      | $$$$$$$$| $$  | $$|  $$$$$$/| $$  | $$
 \______/ |__/      |________/|__/  |__/ \______/ |__/  |__/



*/

#If WinActive("myCommands.ahk ahk_class AutoHotkeyGUI") || WinActive("functionKeys ahk_class AutoHotkeyGUI")
+Enter::Send {asc 010}
#IfWinActive

#IfWinActive, ahk_class AutoHotkeyGUI
Browser_Back::Esc
#IfWinActive


/*
 /$$      /$$ /$$$$$$  /$$$$$$   /$$$$$$
| $$$    /$$$|_  $$_/ /$$__  $$ /$$__  $$
| $$$$  /$$$$  | $$  | $$  \__/| $$  \__/
| $$ $$/$$ $$  | $$  |  $$$$$$ | $$
| $$  $$$| $$  | $$   \____  $$| $$
| $$\  $ | $$  | $$   /$$  \ $$| $$    $$
| $$ \/  | $$ /$$$$$$|  $$$$$$/|  $$$$$$/
|__/     |__/|______/ \______/  \______/


*/


#IfWinNotActive, ahk_class OpusApp 
^Backspace:: Send ^+{Left}{Backspace}
^Delete:: Send ^+{Right}{Delete}
#IfWinNotActive

^!k::
	splashText("Kill what?")
	; InputBox, IM, TaskKill, Kill what?
	if input <> ; if input is not empty
	{
		if input = x
		{
			GroupAdd, WindowsExplorer, ahk_class CabinetWClass 
			WinClose, ahk_group WindowsExplorer
			TrayTip, ,Windows Explorer killed

		}
		Else
		{
			Run, taskkill /F /IM %input%*,,Hide
			TrayTip, ,%input%* killed
		}
	}
	Else
		TrayTip, taskkill, User canceled
	Return

; ::fff::{Home}
; ::jjj::{End}
+^z:: ; Undo  
	loop 5
	{
		Send, ^z
		splashNotify("Undo x " A_Index, 1000) 
	}
	Return
!^z:: ; Undo  
	loop 10
	{
	 Send, ^z
		splashNotify("Undo x " A_Index, 1000) 
	}
	Return
+!^z:: ; Undo  
	loop 20
	{
		Send, ^z
		splashNotify("Undo x " A_Index, 1000) 
	}
	Return


^F9:: ; Sort comma-separated, single line list alphabetically
	clipboard = ; Clear clipboard
	Send, ^c
	ClipWait
	StringReplace, clipboard, clipboard, `,%A_Space%,`,, ALL ; Replace ", " with ","
	StringReplace, clipboard, clipboard, %A_Space%`,,`,, ALL ; Replace " ," with ","
	Sort, clipboard, D, ; Sort %clipboard% alphabetically with comma as delimiter
	StringReplace, clipboard, clipboard, `,,`,%A_Space%, ALL ; Replace "," with ", "
	Send, %clipboard% ; Paste clipboard
	Return

/*
~LAlt::
	if (A_PriorHotkey <> "~LAlt" or A_TimeSincePriorHotkey > 400)
	{
		KeyWait, LAlt
		Return
	}
	openChrome()
	Return
*/



; ^!PrintScreen:: ; Renew IP address
; 	run %ComSpec% ipconfig /renew
; 	TrayTip, ipconfig, renewed
; 	Return

^!F8:: ; Ennumerate highlighted text
	Clipboard :=
		Send ^x
		ClipWait
		Loop, Parse, Clipboard,`r
		{
				line = %A_LoopField%
				StringReplace, line, line,`r,, All
				StringReplace, line, line,`n,, All
				send %A_Index%. %line%`r
		}
		Send {BackSpace}
		Return

^!F9:: ; Lowercase highlighted text
	StringLower Clipboard, Clipboard
	Send %clipboard%
	Return



/*
	/$$$$$$  /$$
 /$$__  $$| $$
| $$  \__/| $$$$$$$   /$$$$$$   /$$$$$$  /$$$$$$/$$$$   /$$$$$$
| $$      | $$__  $$ /$$__  $$ /$$__  $$| $$_  $$_  $$ /$$__  $$
| $$      | $$  \ $$| $$  \__/| $$  \ $$| $$ \ $$ \ $$| $$$$$$$$
| $$    $$| $$  | $$| $$      | $$  | $$| $$ | $$ | $$| $$_____/
|  $$$$$$/| $$  | $$| $$      |  $$$$$$/| $$ | $$ | $$|  $$$$$$$
 \______/ |__/  |__/|__/       \______/ |__/ |__/ |__/ \_______/

=== Chrome === 

*/



/*
 /$$$$$$$$                     /$$
| $$_____/                    | $$
| $$       /$$   /$$  /$$$$$$ | $$  /$$$$$$   /$$$$$$   /$$$$$$   /$$$$$$
| $$$$$   |  $$ /$$/ /$$__  $$| $$ /$$__  $$ /$$__  $$ /$$__  $$ /$$__  $$
| $$__/    \  $$$$/ | $$  \ $$| $$| $$  \ $$| $$  \__/| $$$$$$$$| $$  \__/
| $$        >$$  $$ | $$  | $$| $$| $$  | $$| $$      | $$_____/| $$
| $$$$$$$$ /$$/\  $$| $$$$$$$/| $$|  $$$$$$/| $$      |  $$$$$$$| $$
|________/|__/  \__/| $$____/ |__/ \______/ |__/       \_______/|__/
										| $$
										| $$
										|__/
=== Explorer ===
*/


#e::run %user%\AppData\Roaming\Microsoft\Windows\Recent

#IfWinActive, ahk_class CabinetWClass
!d::
	Send, {F4}
	Send, +{Home}{Delete}
	Return 
; +PgUp::
; 	Send, !{Up}
; 	Send, {Up}{Enter}
; 	Return
; +PgDn::
; 	Send, !{Up}
; 	Send, {Down}{Enter}
; 	Return
#IfWinNotActive

#IfWinActive, ahk_class CabinetWClass, .zip ; When zip file opened in Windows Explorer, press Enter to unzip
{
	Enter::
		Send, {F4}
		Sleep, 200
		Send, ^a^c
		Send, !f{Enter}!{F4}
		Sleep, 500
		Send, {Enter}
		Return
}
#IfWinActive


; ## TOGGLE HIDDEN FILES
+#h::
RegRead, HiddenFiles_Status, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced, Hidden
If HiddenFiles_Status = 2
RegWrite, REG_DWORD, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced, Hidden, 1
Else
RegWrite, REG_DWORD, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced, Hidden, 2
send, {F5}
TrayTip, AutoHotkey, Hidden files toggled
	Return

/*
 /$$$$$$$                                             /$$$$$$$           /$$             /$$    
| $$__  $$                                           | $$__  $$         |__/            | $$    
| $$  \ $$ /$$$$$$  /$$  /$$  /$$  /$$$$$$   /$$$$$$ | $$  \ $$ /$$$$$$  /$$ /$$$$$$$  /$$$$$$  
| $$$$$$$//$$__  $$| $$ | $$ | $$ /$$__  $$ /$$__  $$| $$$$$$$//$$__  $$| $$| $$__  $$|_  $$_/  
| $$____/| $$  \ $$| $$ | $$ | $$| $$$$$$$$| $$  \__/| $$____/| $$  \ $$| $$| $$  \ $$  | $$    
| $$     | $$  | $$| $$ | $$ | $$| $$_____/| $$      | $$     | $$  | $$| $$| $$  | $$  | $$ /$$
| $$     |  $$$$$$/|  $$$$$/$$$$/|  $$$$$$$| $$      | $$     |  $$$$$$/| $$| $$  | $$  |  $$$$/
|__/      \______/  \_____/\___/  \_______/|__/      |__/      \______/ |__/|__/  |__/   \___/  
																																																
																																																
																																																
*/
#IfWinActive, ahk_class PPTFrameClass
; ^i::^k ; Remap keyboard shortcuts in Danish version to international 
; ^k::^i
; ^f::^b
; ^b::^f
#IfWinActive


/*
 /$$$$$$$$                               /$$
| $$_____/                              | $$
| $$       /$$   /$$  /$$$$$$$  /$$$$$$ | $$
| $$$$$   |  $$ /$$/ /$$_____/ /$$__  $$| $$
| $$__/    \  $$$$/ | $$      | $$$$$$$$| $$
| $$        >$$  $$ | $$      | $$_____/| $$
| $$$$$$$$ /$$/\  $$|  $$$$$$$|  $$$$$$$| $$
|________/|__/  \__/ \_______/ \_______/|__/


=== Excel === 

*/

#IfWinActive, ahk_class XLMAIN
; ^i::^k ; Remap keyboard shortcuts in Danish version to international 
; ^k::^i
; ^f::^b
; ^b::^f

+!Left:: ; Hide active column 
	xl := ComObjActive("Excel.Application") 
	xl.Selection.EntireColumn.Hidden := True ; !xl.ActiveCell.EntireColumn.Hidden
	Return 

+!Right:: ; Unhide active column 
	xl := ComObjActive("Excel.Application") 
	xl.Selection.EntireColumn.Hidden := False ; !xl.ActiveCell.EntireColumn.Hidden
	Return 


+!Up:: ; Hide active row 
	xl := ComObjActive("Excel.Application") 
	xl.Selection.EntireRow.Hidden := True ; !xl.ActiveCell.EntireColumn.Hidden
	Return 

+!Down:: ; Unhide active row 
	xl := ComObjActive("Excel.Application") 
	xl.Selection.EntireRow.Hidden := False ; !xl.ActiveCell.EntireColumn.Hidden
	Return 



; +^h:: ; Toggle hide selected column 
; 	xl := ComObjActive("Excel.Application") 
; 	xl.ActiveCell.EntireRow.Select
; 	; xl.Application.Selection.EntireColumn.Hidden := !xl.Application.Selection.EntireColumn.Hidden
; 	Return 

^SC00C::comObjActive("Excel.Application").Application.ActiveWindow.Zoom += 10 ; Zoom in
^SC035::comObjActive("Excel.Application").Application.ActiveWindow.Zoom -= 10 ; Zoom out
^0::comObjActive("Excel.Application").Application.ActiveWindow.Zoom := 100 ; Zoom reset 

#IfWinActive

+!x:: ; Send highlighted text to new Excel workbook
	clipboard =
	Send, ^c
	ClipWait
	clipboard = %clipboard%
	xl := ComObjCreate("Excel.Application")
	xl.Visible := True ;by default excel sheets are invisible
	xl.Workbooks.Add ;add a new workbook
	WinWaitActive, ahk_class XLMAIN
	Send, ^v
	Return
:*:\groups::
	; Add Shapes and text to the active Word document from Excel data
	xl := ComObjActive("Excel.Application")
	Loop % xl.ActiveSheet.UsedRange.Columns.Count ; Loop through used columns
	{
		; And add one shape for each columns
		oWord := ComObjActive("Word.Application")
		x := 100 + (20 * A_Index)
		y := 100 + (20 * A_Index)
		oWord.ActiveDocument.Shapes.AddShape(5, x, y, 150, 200)
		oWord.ActiveDocument.Shapes(oWord.ActiveDocument.Shapes.Count).Select
		oWord.Application.Selection.ShapeRange.WrapFormat.Type := 0
	}
	Loop % xl.ActiveSheet.UsedRange.Columns.Count ; Loop through used columns
	{
		; And populate each shape with its corresponding excel data
		xl.ActiveSheet.Columns(A_Index).Select
		xl.Selection.Copy
		clipboard = %clipboard%
		oWord.ActiveDocument.Shapes(A_Index).Select
		oWord.Selection.Paste
	}
	Return

/*
 /$$      /$$                           /$$
| $$  /$ | $$                          | $$
| $$ /$$$| $$  /$$$$$$   /$$$$$$   /$$$$$$$
| $$/$$ $$ $$ /$$__  $$ /$$__  $$ /$$__  $$
| $$$$_  $$$$| $$  \ $$| $$  \__/| $$  | $$
| $$$/ \  $$$| $$  | $$| $$      | $$  | $$
| $$/   \  $$|  $$$$$$/| $$      |  $$$$$$$
|__/     \__/ \______/ |__/       \_______/

=== Word ===

*/



#IfWinActive, ahk_class OpusApp 

+!^0::ComObjActive("Word.Application").Selection.Style := ComObjActive("Word.Application").ActiveDocument.Styles("Normal") ; Apply "Normal" style


!Space::ComObjActive("Word.Application").Selection.Words(1).Select ; Select the current word

+^Space:: ; Select sentence 
	oWord := ComObjActive("Word.Application")
	oWord.Selection.MoveLeft(Unit:=3)  ; Move to start of sentence
	oWord.Selection.ExtendMode := True ; Turns on extend mode 
	oWord.Selection.MoveRight(Unit:=3) ; Expand selection to whole sentence
	oWord.Selection.ExtendMode := False ; Turns off extend mode 
	Return 
	
+^F7::ComObjActive("Word.Application").Selection.LanguageID := 1024 ; No proofing
+F9::ComObjActive("Word.Application").Selection.Sort 

+F11:: ; Full screen 
	aW := ComObjActive("Word.Application").ActiveWindow
	aW.View.FullScreen := ! aW.View.FullScreen
	Return 

F11:: ; Distraction free 
	WinMaximize, A
	ComObjActive("Word.Application").Application.ScreenUpdating := True

	aW := ComObjActive("Word.Application").ActiveWindow
	aW.View.Type := 3 ; wdPrintView = 3
	aW.View.DisplayPageBoundaries := !aW.View.DisplayPageBoundaries
	aW.DocumentMap := "False" ; !aW.DocumentMap
	aW.ActivePane.DisplayRulers := "False" ; !aW.ActivePane.DisplayRulers
	if aW.UsableHeight < 550 ; The UsableHeight is greater when the ribbon is hidden
		aW.ToggleRibbon
	Return


!Up::ComObjActive("Word.Application").ActiveDocument.ActiveWindow.SmallScroll(0,1,0,0)
!Down::ComObjActive("Word.Application").ActiveDocument.ActiveWindow.SmallScroll(1,0,0,0)
!Left::ComObjActive("Word.Application").ActiveDocument.ActiveWindow.SmallScroll(0,0,0,1)
!Right::ComObjActive("Word.Application").ActiveDocument.ActiveWindow.SmallScroll(0,0,1,0)

^!SC00C::ComObjActive("Word.Application").Selection.Font.Grow
^!SC035::ComObjActive("Word.Application").Selection.Font.Shrink

^!b:: 
	toggle++
	if  mod(toggle, 2) = 1
		ComObjActive("Word.Application").Selection.Font.Subscript := 1
	if  mod(toggle, 2) = 0
		ComObjActive("Word.Application").Selection.Font.Subscript := 0
	Return

^!p:: 
	toggle++
	if  mod(toggle, 2) = 1
		ComObjActive("Word.Application").Selection.Font.Superscript := 1
	if  mod(toggle, 2) = 0
		ComObjActive("Word.Application").Selection.Font.Superscript := 0
	Return

+^F4::
	Try
		ComObjActive("Word.Application").Documents.Close
	WinClose, ahk_class OpusApp
	Return 
; ^i::^k ; Remap keyboard shortcuts in Danish version to international 
; ^k::^i
; ^f::^b
; ^b::^f

Ins::
+^v:: ; Paste as plain text 
	Clipboard = %Clipboard% ; Trim leading and trailing spaces. Remove formatting. 
	Send, ^v
	Return

Browser_Back::Esc

; !+e::editComment() ; Edit last comment

	
^0:: ; Reset zoom
	oWord := ComObjActive("Word.Application")
	oWord.ActiveDocument.ActiveWindow.View.Zoom.Percentage := 100
	Return
^SC00C:: ; Zoom in
	oWord := ComObjActive("Word.Application")
	myZoom := oWord.ActiveDocument.ActiveWindow.View.Zoom.Percentage
	oWord.ActiveDocument.ActiveWindow.View.Zoom.Percentage := myZoom + 5
		Return
^SC035:: ; Zoom out
	oWord := ComObjActive("Word.Application")
	myZoom := oWord.ActiveDocument.ActiveWindow.View.Zoom.Percentage
	oWord.ActiveDocument.ActiveWindow.View.Zoom.Percentage := myZoom - 5
		Return
^+i:: ; Insert file
	FileSelectFile, file
	Loop, %file%
		extension := A_LoopFileExt
	if ErrorLevel <> 1
	{
		if extension = png
			ComObjActive("Word.Application").Application.Selection.InlineShapes.AddPicture(file)
		Else
			ComObjActive("Word.Application").Application.Selection.InsertFile(file)	
	}
	Return

; ^Space::ComObjActive("Word.Application").Selection.Tables(1).Select

+!Down:: ; From current position, expand selection sentence-by-sentence
	oWord := ComObjActive("Word.Application")
		oWord.Selection.ExtendMode := True ; Turns on extend mode 
		oWord.Selection.MoveRight(Unit:=3) ; Expand selection to whole sentence
		oWord.Selection.ExtendMode := False ; Turns off extend mode 
		Return

+!Up:: ; From current position, expand selection sentence-by-sentence
	oWord := ComObjActive("Word.Application")
		oWord.Selection.ExtendMode := True ; Turns on extend mode
		oWord.Selection.MoveLeft(Unit:=3) ; Expand selection to whole sentence
		oWord.Selection.ExtendMode := False ; Turns off extend mode 
		Return

+!Right::ComObjActive("Word.Application").Selection.MoveRight(Unit:=3)  ; Move to end of sentence
+!Left::ComObjActive("Word.Application").Selection.MoveLeft(Unit:=3)  ; Move to start of sentence

; +^!b:: send !nck ; Draw textbox
; ^!b::  ; Insert highlighted text in simple textbox
; 	Send ^x
; 	Sleep 50
; 	Send !nc{Enter}
; 	Sleep 1000
; 	Send ^v{Esc}{Esc}
; 	Return
; Square wrap
; ^!w:: send !jof
; Set margins
; ^!3:: send !ajb

!^BackSpace::
	try
		ComObjActive("Word.Application").Selection.Rows.Delete
	Return


#IfWinActive

/*
	/$$$$$$            /$$       /$$ /$$                               /$$$$$$$$                    /$$
| $$  \__/ /$$   /$$| $$$$$$$ | $$ /$$ /$$$$$$/$$$$   /$$$$$$          | $$  /$$$$$$  /$$   /$$ /$$$$$$
|  $$$$$$ | $$  | $$| $$__  $$| $$| $$| $$_  $$_  $$ /$$__  $$         | $$ /$$__  $$|  $$ /$$/|_  $$_/
 \____  $$| $$  | $$| $$  \ $$| $$| $$| $$ \ $$ \ $$| $$$$$$$$         | $$| $$$$$$$$ \  $$$$/   | $$
 /$$  \ $$| $$  | $$| $$  | $$| $$| $$| $$ | $$ | $$| $$_____/         | $$| $$_____/  >$$  $$   | $$ /$$
|  $$$$$$/|  $$$$$$/| $$$$$$$/| $$| $$| $$ | $$ | $$|  $$$$$$$         | $$|  $$$$$$$ /$$/\  $$  |  $$$$/
 \______/  \______/ |_______/ |__/|__/|__/ |__/ |__/ \_______/         |__/ \_______/|__/  \__/   \___/



=== Sublime Text === 
*/

#IfWinActive, ahk_class PX_WINDOW_CLASS
!F4::!F3
#IfWinActive


!^s:: ; Copy highlighted text to new tab in SublimeText
	Send ^c 
	ClipWait
    IfWinExist, ahk_class PX_WINDOW_CLASS
        WinActivate
    Else
        run %onedrive%\Apps\Sublime Text Build 3143 x64\sublime_text.exe
    WinWaitActive, ahk_class PX_WINDOW_CLASS
	Sleep 200
	Send ^n
	Sleep 200
	Send ^v
	Return

; ^!Tab::Send, {AppsKey}as ; AlignTab (by spaces)

/*
 /$$      /$$
| $$$    /$$$
| $$$$  /$$$$  /$$$$$$  /$$   /$$  /$$$$$$$  /$$$$$$
| $$ $$/$$ $$ /$$__  $$| $$  | $$ /$$_____/ /$$__  $$
| $$  $$$| $$| $$  \ $$| $$  | $$|  $$$$$$ | $$$$$$$$
| $$\  $ | $$| $$  | $$| $$  | $$ \____  $$| $$_____/
| $$ \/  | $$|  $$$$$$/|  $$$$$$/ /$$$$$$$/|  $$$$$$$
|__/     |__/ \______/  \______/ |_______/  \_______/

=== Mouse ===

*/

MButton::AppsKey
~RButton & WheelDown::Send ^{Tab} ; Change tabs
~RButton & WheelUp::Send ^+{Tab} ; Change tabs
!WheelDown::Send ^{Tab} ; Change tabs
!WheelUp::Send ^+{Tab} ; Change tabs
; ~LButton & RButton::SendInput {LWin down}{Tab}{LWin up} ; Send {Browser_Back} 
; ~RButton::
; 	if (A_PriorHotkey <> "~RButton" or A_TimeSincePriorHotkey > 400)
; 	{
; 			; Too much time between presses, so this isn't a double-press.
; 			KeyWait, RButton
; 			return
; 	}
; 	Send {Browser_Forward} 
; 	Return

#IfWinActive, ahk_class Chrome_WidgetWin_1
+WheelLeft::
	Send {Browser_Back}
	Sleep 750 
	Return
+WheelRight::
	Send {Browser_Forward}
	Sleep 750
	Return 
#IfWinActive 
; RButton & WheelDown:: Send {XButton1} ; the 'back' button on a mouse
; RButton & WheelUp:: Send {XButton2} ; the 'forward' button on a mouse
; !LButton:: Send {XButton1} ; the 'back' button on a mouse
; !RButton:: Send {XButton2} ; the 'forward' button on a mouse

/*
 /$$   /$$                     /$$                                           /$$
| $$  /$$/                    | $$                                          | $$
| $$ /$$/   /$$$$$$  /$$   /$$| $$$$$$$   /$$$$$$   /$$$$$$   /$$$$$$   /$$$$$$$
| $$$$$/   /$$__  $$| $$  | $$| $$__  $$ /$$__  $$ |____  $$ /$$__  $$ /$$__  $$
| $$  $$  | $$$$$$$$| $$  | $$| $$  \ $$| $$  \ $$  /$$$$$$$| $$  \__/| $$  | $$
| $$\  $$ | $$_____/| $$  | $$| $$  | $$| $$  | $$ /$$__  $$| $$      | $$  | $$
| $$ \  $$|  $$$$$$$|  $$$$$$$| $$$$$$$/|  $$$$$$/|  $$$$$$$| $$      |  $$$$$$$
|__/  \__/ \_______/ \____  $$|_______/  \______/  \_______/|__/       \_______/
										 /$$  | $$
										|  $$$$$$/

										 \______/
=== Keyboard ===
*/


#If GetKeyState("CapsLock", "T")
~Tab::Down 
#If 


/*
 /$$      /$$ /$$                 /$$
| $$  /$ | $$|__/                | $$
| $$ /$$$| $$ /$$ /$$$$$$$   /$$$$$$$  /$$$$$$  /$$  /$$  /$$  /$$$$$$$
| $$/$$ $$ $$| $$| $$__  $$ /$$__  $$ /$$__  $$| $$ | $$ | $$ /$$_____/
| $$$$_  $$$$| $$| $$  \ $$| $$  | $$| $$  \ $$| $$ | $$ | $$|  $$$$$$
| $$$/ \  $$$| $$| $$  | $$| $$  | $$| $$  | $$| $$ | $$ | $$ \____  $$
| $$/   \  $$| $$| $$  | $$|  $$$$$$$|  $$$$$$/|  $$$$$/$$$$/ /$$$$$$$/
|__/     \__/|__/|__/  |__/ \_______/ \______/  \_____/\___/ |_______/


=== Windows === 
*/
; #IfWinActive, ahk_class #32770
; ; #If WinActive("Gem som ahk_class #32770") || WinActive("Save File As ahk_class #32770")
; Tab::Send, {Down}
; <::\
; #IfWinActive

+!^b::WinSet, Style, ^0xC00000, A
+!^w::WinSet, Style, ^0x840000, A

#IfWinNotActive, ahk_class OpusApp ;  ahk_group, Office
^!Esc:: ; Toggle hide Windows taskbar
	toggle++
	if  mod(toggle, 2) = 1
	{
		hideTaskbar(1)
	}
	if  mod(toggle, 2) = 0
	{
		hideTaskbar(0)
	}
	Return
#IfWinNotActive


; #Left:: ; WindowPad toggle
; 	winToggle++
; 	if  mod(winToggle, 2) = 1
; 		WinMove, A,, 0,28, 683, 768
; 	Else
; 		WinMove, A,, 0,28, 342, 768
; 	Return

; #PgUp::WinMove, A,, -8, 29, 1382, 746 ; Maximize

+#PgDn:: ; Extra WindowPad toggle
	winToggle++
	if  mod(winToggle, 2) = 1
		WinMove, A,, 342,, 1024,
	Else
		WinMove, A,, 342,, 683,
	Return

+#PgUp:: ; Extra WindowPad toggle
	winToggle++
	if  mod(winToggle, 2) = 1
		WinMove, A,, 220, 28, 1146, 617
	Else
		WinMove, A,, -6, 24, 1378, 751
	Return

^!Space:: ; Always on top
WinGet, currentWindow, ID, A
WinGet, ExStyle, ExStyle, ahk_id %currentWindow%
if (ExStyle & 0x8)  ; 0x8 is WS_EX_TOPMOST.
{
	Winset, AlwaysOnTop, off, ahk_id %currentWindow%

	SplashImage,, x0 y0 b fs12, OFF always on top.
	Sleep, 1500
	SplashImage, Off
}
else
{
	WinSet, AlwaysOnTop, on, ahk_id %currentWindow%
	SplashImage,,x0 y0 b fs12, ON always on top.
	Sleep, 1500
	SplashImage, Off
}
return

+^!-:: ; Decrese transparency
WinGet, currentWindow, ID, A
if not (%currentWindow%)
{
	%currentWindow% := 255
}
if (%currentWindow% != 255)
{
	%currentWindow% += 5
	WinSet, Transparent, % %currentWindow%, ahk_id %currentWindow%
}
SplashImage,,w100 x0 y0 b fs12, % %currentWindow%
SetTimer, TurnOffSI, 1000, On
	Return

+^!+:: ; Increase transparency
SplashImage, Off
WinGet, currentWindow, ID, A
if not (%currentWindow%)
{
	%currentWindow% := 255
}
if (%currentWindow% != 5)
{
	%currentWindow% -= 5
	WinSet, Transparent, % %currentWindow%, ahk_id %currentWindow%
}
SplashImage,, w100 x0 y0 b fs12, % %currentWindow%
SetTimer, TurnOffSI, 1000, On
	Return

TurnOffSI:
SplashImage, off
SetTimer, TurnOffSI, 1000, Off
	Return

