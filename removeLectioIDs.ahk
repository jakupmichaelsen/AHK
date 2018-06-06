#NoTrayIcon
#SingleInstance, force 
IniRead()
Progress, CT%header_color% CW%background% B1 Y0 ZH0 CTFF0000, Navigate to relevant folder and press F1 to remove Lectio filename prefixes (ESC to cancel)

; Progress, show CT%header_color% CW%background% fm20 WS700 c00 x0 w300 h125 zy60 zh0 B0, , %text%,, %header% 
Esc::ExitApp
F1::
folder := Explorer_GetPath()
; WinGetTitle, folder, A
Loop, %folder%\*.*
	fileCount := A_Index

Loop, %folder%\*.*
{
	if A_Index > %fileCount% 
		ExitApp
	SplitPath, A_LoopFileFullPath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
	oldFilename := folder . "\" . OutFileName
	StringTrimLeft, OutFileName, OutFileName, 12
	newFilename := folder . "\" . OutFileName
	FileMove, %oldFilename%, %newFilename%, 1
}
ExitApp

