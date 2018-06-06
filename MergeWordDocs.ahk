#SingleInstance, force


FileSelectFile, files, M 

Loop, Parse, files 
{
	MsgBox, %A_Fieldname%
}
; ComObjActive("Word.Application").Application.Selection.InsertFile(file)
