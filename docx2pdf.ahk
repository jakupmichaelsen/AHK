global oWord := ComObjCreate("Word.Application") 


FileSelectFolder, Folder, , , Select the root folder containing your Word documents
FileCreateDir, %Folder%\pdf
mergeDocs(Folder)
oWord.quit
Msgbox Done!
Run, %Folder%
ExitApp, 

mergeDocs(Folder)
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
	 		ToolTip, Converting %A_LoopFileName%, 0, 0
	 		oWord.Application.Documents.Open(A_LoopFileFullPath)
	 		SplitPath, A_LoopFileFullPath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
		 	oWord.Application.ActiveDocument.SaveAs2(Folder . "\pdf\" . OutNameNoExt . ".pdf", 17) 
			oWord.Application.ActiveDocument.Close()
		}
	Sleep 5000
 	}
}

