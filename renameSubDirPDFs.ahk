#SingleInstance, force 
Progress, CT%header_color% CW%background% B1 Y0 ZH0 CTFF0000, Navigate to relevant folder and press F1 to rename subdirectory PDFs by their folders and move to current root directory (ESC to cancel)

Esc::ExitApp

F1::
	rootDir := Explorer_GetPath()
	Loop Files, %rootDir%\*.pdf, R  ; Recurse into subfolders.
	{
		fileName := StrReplace(A_LoopFileDir, rootDir) ; Get the file name 
		StringTrimLeft, fileName, fileName, 2
 		ToolTip, Converting %A_LoopFileName%, 0, 0
		FileMove, %A_LoopFileFullPath%, %rootDir%\%fileName%.pdf
	}
	Msgbox Done!
	ExitApp

	; MsgBox, 4, %A_ScriptTitle%, Would you like to convert PDFs to docx?
	; IfMsgBox, Yes
	; {
	; 	Run pdf2docx.ahk 
	; 	ExitApp
	; }
	; Else 
	; 	ExitApp
