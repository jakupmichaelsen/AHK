#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
#NoTrayIcon


SendMode Input
TrayTip, %A_ScriptName%, Script loaded
AutoReload()
SetTitleMatchMode, 2

#IfWinActive, WorkFlowy
; ~!Left::Browser_Back
; ~!Right::Browser_Forward
; !PgUp::!Left
; !PgDn::!Right
::'::’
::"::”
+^v::
	StringReplace, Clipboard, Clipboard, https://
	StringReplace, Clipboard, Clipboard, http://
	Send ^v 
	Return 

wStringInfo:
	Global rChar
	Global wLen 
	Global isComma
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Send, ^c 
	ClipWait 
	StringRight, rChar, Clipboard, 1
	StringLen, wLen, Clipboard 
	if rChar = ,
		isComma = 1
	Else 
		isComma = 0
	Clipboard = %ClipSaved%  ; Restore Clipboard
	Return 

+^h::
;	Send, ^{Left}
;	Send, +^{Right}
;	Send, +{Left}
;	Gosub, wStringInfo
;	if isComma
;		Send, +{Left}
	Send, ^b
	Send, ^u
	Send, ^i
	Send, {Right}
	Return 

+PgDn::+!0
+PgUp::+!9
^Up::
	Send ^a
	Send {Left}
	return
^Down::
	Send ^a
	Send {Right}
	return
+F2::Click 3
; +Esc::
;     splashText("WorkFlowy query")
; 	if input <>
;     run https://workflowy.com/#/?q=%input%
Return 
^!F9:: ; Ennumerate highlighted text
    Clipboard :=
    Send ^x
    ClipWait
    Send {Enter}{Tab}
    Loop, Parse, Clipboard,`r
    {
        line = %A_LoopField%
        StringReplace, line, line,`r,, All
        StringReplace, line, line,`n,, All
        send %A_Index%. %line%`r
    }
    Send {BackSpace}
    Return

#IfWinActive