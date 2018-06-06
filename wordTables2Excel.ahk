#SingleInstance, force
AutoReload()
TrayTip, %A_ScriptName%, Loaded
; 1 table per file
oWord := ComObjActive("Word.Application")
xl := ComObjActive("Excel.Application") 

F3::
folder := Explorer_GetPath()
Loop, %folder%\*.*
{
	if A_LoopFileAttrib contains H,R,S  ; Skip any file that is either Hidden, Read-only, or System
    	continue  
	
	oWord.Application.Documents.Open(A_LoopFileFullPath)
	oWord.Selection.Tables(1).Select
	oWord.Selection.Copy
	WinActivate, ahk_class XLMAIN
	WinWaitActive, ahk_class XLMAIN
	Send, ^{End}
	Send, ^{Left}
	Send ^v 
}
