#SingleInstance, force 
global oWord := ComObjCreate("Word.Application") 

Progress, CT%header_color% CW%background% B1 Y0 ZH0 CTFF0000, Navigate to relevant folder and press F1 to insert feedbackTemplate at the start of each docx (ESC to cancel)

; try
; 	global oWord := ComObjActive("Word.Application")
; catch
; {
; 	MsgBox,16, %A_ScriptName%, No active MS Word document found.`rExiting.
; 	ExitApp
; }


IniRead, onedrive, splashUI.ini, paths, onedrive


global feedbackTemplate
feedbackTemplate = %onedrive%\Adjunkt\Grading\Templates\exam_DigiA.rtf

Esc::ExitApp
F1::
	rootDir := Explorer_GetPath()

	Loop Files, %rootDir%\*.docx
	{
		SplitPath, A_LoopFileFullPath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
		oWord.Application.Documents.Open(feedbackTemplate)
		; oWord.Application.Visible := True 
		; WinActivate, ahk_class OpusApp
		; WinWaitActive, ahk_class OpusApp	
		splashNotify("Processing " . OutFileName) 
		oWord.Selection.EndKey(6) ; wdStory = 6
		; Sleep 200
		; splashNotify("Inserting break") 
		oWord.Selection.InsertBreak(2) ; wdSectionBreakNextPage
		; Sleep 200
		oWord.Application.Selection.InsertFile(A_LoopFileFullPath)
		; splashNotify("Inserting file") 
		; Sleep 200
	 	oWord.Application.ActiveDocument.SaveAs2(A_LoopFileFullPath, 16) 
		; splashNotify("Saving...") 
		; Sleep 200
		oWord.Application.ActiveDocument.Close()
	}
	oWord.Application.Quit
	MsgBox Done!
	ExitApp	