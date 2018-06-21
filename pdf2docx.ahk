global oWord := ComObjCreate("Word.Application") 


Progress, CT%header_color% CW%background% B1 Y0 ZH0 CTFF0000, Navigate to the root folder containing your PDFs and press F1 to convert to DOCX (ESC to cancel)

Esc::ExitApp
F1::
	Folder := Explorer_GetPath()
	FileCreateDir, %Folder%\docx
	FileCreateDir, %Folder%\pdf
	convert(Folder)
	oWord.quit
	Msgbox Done!
	ExitApp

convert(Folder)
{
	global
	loop, %Folder%\*.pdf
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
		 	oWord.Application.ActiveDocument.SaveAs2(Folder . "\docx\" . OutNameNoExt . ".docx", 16) 
			oWord.Application.ActiveDocument.Close()
			FileMove, %A_LoopFileFullPath%, %Folder%\pdf
		}
 	}
}

