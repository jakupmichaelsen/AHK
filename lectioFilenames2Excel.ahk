#SingleInstance, force
Progress, CT%header_color% CW%background% B1 Y0 ZH0 CTFF0000, Navigate to relevant folder and press F1 to copy filenames to 'Filenames.xlsx' (ESC to cancel)
xl := ComObjCreate("Excel.Application") 

xl.Application.Workbooks.Add() 

Esc::
	xl.Quit()
	xl := 
	ExitApp
F1::
folder := Explorer_GetPath()
Loop, %folder%\*.*
{
	SplitPath, A_LoopFileFullPath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
	xl.Range("B" . A_Index).Value := OutFileName
	StringTrimLeft, trimmedFileName, OutFilename, 12
	xl.Range("A" . A_Index).Value := trimmedFileName
}
xl.Columns("A:B").AutoFit
path = %folder%\Filenames.xlsx
xl.ActiveWorkbook.SaveAs(path)
xl.Quit
xl := ""

ExitApp