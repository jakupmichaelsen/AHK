; Merge Word files 
global oWord := ComObjCreate("Word.Application") 
oWord.Application.Visible := False 
FileSelectFolder, Folder, 
loadDocs(Folder)
Msgbox Done!
return

loadDocs(Folder)
{
	loop, %Folder%\*.docx
 	{
	Document := oWord.Application.Documents.Open(A_LoopFileFullPath)
	ToolTip, Converting %A_LoopFileName%, 
	Document.SaveAs2("\" . A_LoopFileName . ".pdf", 17) 
	Document.Close()    
	}
}
