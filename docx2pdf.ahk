global oWord := ComObjCreate("Word.Application") 


Progress, CT%header_color% CW%background% B1 Y0 ZH0 CTFF0000, Navigate to the root folder containing your DOCXs and press F1 to convert to PDF (ESC to cancel)

Esc::ExitApp
F1::
	Folder := Explorer_GetPath()
	FileCreateDir, %Folder%\pdf
	convert(Folder)
	oWord.quit
	Msgbox Done!
	ExitApp

convert(Folder)
{
	global
	loop, %Folder%\*.docx
 	{
 		docIndex = %A_Index%
 		StringGetPos, pos, A_LoopFileName, ~
 		if pos = 0
  			Continue
  		Else
  		{
			splashNotify("Converting: " . A_LoopFileName) 
	 		oWord.Application.Documents.Open(A_LoopFileFullPath)
	 		SplitPath, A_LoopFileFullPath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
		 	oWord.Application.ActiveDocument.SaveAs2(Folder . "\pdf\" . OutNameNoExt . ".pdf", 17) 
			oWord.Application.ActiveDocument.Close()
		}
 	}
}

