oWord := ComObjActive("Word.Application") 
tCount := oWord.ActiveDocument.Tables.Count

MsgBox,1,%A_ScriptTitle%, Deleting %tCount% tables 
IfMsgBox, Ok
{
	while % oWord.ActiveDocument.Tables.Count <> 0
		oWord.ActiveDocument.Tables(1).Delete
}
IfMsgBox, Cancel
	ExitApp

ExitApp