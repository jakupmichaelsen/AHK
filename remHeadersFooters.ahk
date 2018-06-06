oWord := ComObjActive("Word.Application")
SetTimer, ChangeButtonNames, 50 
MsgBox, 3, Add or Delete, Choose a button:
IfMsgBox, Yes
	oWord.ActiveDocument.Sections(1).Headers(1).Range.Delete
IfMsgBox, No
	oWord.ActiveDocument.Sections(1).Footers(1).Range.Delete
IfMsgBox, Cancel 
{
	oWord.ActiveDocument.Sections(1).Headers(1).Range.Delete
	oWord.ActiveDocument.Sections(1).Footers(1).Range.Delete
}
return 

ChangeButtonNames: 
IfWinNotExist, Add or Delete
    return  ; Keep waiting.
SetTimer, ChangeButtonNames, off 
WinActivate 
ControlSetText, Button1, &Headers
ControlSetText, Button2, &Footers 
ControlSetText, Button3, &Both
return
