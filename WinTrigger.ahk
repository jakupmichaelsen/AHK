;========================================================================
; 
; Template:     WinTrigger (former OnOpen/OnClose)
; Description:  Act upon (de)activation/(un)existance of programs/windows
; Online Ref.:  http://www.autohotkey.com/forum/viewtopic.php?t=63673
;
; Last Update:  15/Mar/2010 17:30
;
; Created by:   MasterFocus
;               http://www.autohotkey.net/~MasterFocus/AHK/
;
; Thanks to:    Lexikos, for improving it significantly
;               http://www.autohotkey.com/forum/topic43826.html#267338
;
;========================================================================
;
; This template contains two examples by default. You may remove them.
;
; * HOW TO ADD A PROGRAM to be checked upon (de)activation/(un)existance:
;
; 1. Add a variable named ProgWinTitle# (Configuration Section)
; containing the desired title/ahk_class/ahk_id/ahk_group
;
; 2. Add a variable named WinTrigger# (Configuration Section)
; containing the desired trigger ("Exist" or "Active")
;
; 3. Add labels named LabelTriggerOn# and/or LabelTriggerOff#
; (Custom Labels Section) containing the desired actions
;
; 4. You may also change CheckPeriod value if desired
;
;========================================================================

#Persistent
#SingleInstance, force
#NoTrayIcon
AutoReload()


; ------ ------ CONFIGURATION SECTION ------ ------

; Program Titles

ProgWinTitle1 = FirstClass forbindelse ; FirstClass login
WinTrigger1 = Active ; Exist

ProgWinTitle2 = ahk_class PX_WINDOW_CLASS ; Sublime Text 
WinTrigger2 = Active ; Exist ; Active 

ProgWinTitle3 = Windows PowerShell ahk_class ConsoleWindowClass ; PowerShell
WinTrigger3 = Active ; Exist 

ProgWinTitle4 = Fejl ahk_class #32770 ; FirstClass forbindelse afbrudt 
WinTrigger4 = Active ; Exist 

ProgWinTitle5 = Skrivebord : Silkeborg Gymnasium ahk_class SAWindow ; FirstClass
WinTrigger5 = Active ; Exist 

ProgWinTitle6 = ahk_class _macr_PDAPP_Native_frame_window_CS4
WinTrigger6 = Active ; Exist 

ProgWinTitle7 = ahk_class ConsoleWindowClass
WinTrigger7 = Active ; Exist 

; SetTitleMatchMode, 2
; ProgWinTitle4 = WorkFlowy
; WinTrigger4 = Active ; Exist 
; SetTitleMatchMode, 1


; SetTimer Period
CheckPeriod = 200

; ------ END OF CONFIGURATION SECTION ------ ------

SetTimer, LabelCheckTrigger, %CheckPeriod%
Return

; ------ ------ ------

LabelCheckTrigger:
  While ( ProgWinTitle%A_Index% != "" && WinTrigger := WinTrigger%A_Index% )
    if ( !ProgRunning%A_Index% != !Win%WinTrigger%( ProgWinTitle := ProgWinTitle%A_Index% ) )
      GoSubSafe( "LabelTriggerO" ( (ProgRunning%A_Index% := !ProgRunning%A_Index%) ? "n" : "ff" ) A_Index )
Return

; ------ ------ ------

GoSubSafe(mySub)
{
  if IsLabel(mySub)
    GoSub %mySub%
}

; ------ ------ CUSTOM LABEL SECTION ------ ------

; LabelTriggerOn1: 
;	Send {Enter} 
;	Return
LabelTriggerOff1:
  	;  MsgBox % "A_ThisLabel:`t" A_ThisLabel "`nProgWinTitle:`t" ProgWinTitle "`nWinTrigger:`t" WinTrigger
  	Return
LabelTriggerOn2: ; Sublime Text
	WinSet, Style, -0xC00000, A ; Windowless
	; WinSet, Style, -0x840000, A ; Windowless
	; WinMove, A,, 0, 0, 1926, 1043 ; Maximize
	; WinMaximize
  	Return

LabelTriggerOn3:
	; MsgBox WinSet, Style, -0x840000, A ; Borderless
	WinSet, Style, -0xC00000, A ; Windowless
  	; WinSet, Transparent, 230, A
  	Return 

LabelTriggerOn4:
	WinActivate, Fejl ahk_class #32770
	WinWaitActive, Fejl ahk_class #32770
	Send, {Esc}
	; IfWinExist, ahk_class SAWindow
	; 	WinClose, ahk_class SAWindow
	; Sleep 1000
	; Run, C:\Program Files (x86)\FirstClass\fcc64.exe
	Return

LabelTriggerOn5: ; FirstClass
	WinSet, Style, -0xC00000, A ; Windowless
	WinMaximize
	; WinMove, A,, -9, -9, 1938, 1060 ; Maximize
  	Return 

LabelTriggerOn6:
	MsgBox, [ Options, Title, Text, Timeout]
	Return

LabelTriggerOn7:
	WinSet, Style, -0x840000, A ; Windowless
	Return

; ------ END OF CUSTOM LABEL SECTION ------ ------