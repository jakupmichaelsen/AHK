oWord := ComObjActive("Word.Application")
dName := oWord.ActiveDocument.Name
counter = 0
SetTimer, ChangeButtonNames, 50 

MsgBox, 3, Add or Delete, Choose a button:
IfMsgBox, Yes
	loop % oWord.ActiveDocument.Sections.Count 
	{
		splashNotify("Cleanup in section " . A_Index, 2000) 
		oWord.ActiveDocument.Sections(A_Index).Headers(1).Range.Delete
		counter++
	}
IfMsgBox, No
	loop % oWord.ActiveDocument.Sections.Count 
	{
		splashNotify("Cleanup in section " . A_Index, 2000) 
		oWord.ActiveDocument.Sections(A_Index).Footers(1).Range.Delete
		counter++
	}
IfMsgBox, Cancel 
	loop % oWord.ActiveDocument.Sections.Count 
	{
		splashNotify("Cleanup in section " . A_Index, 2000) 
		oWord.ActiveDocument.Sections(A_Index).Headers(1).Range.Delete
		oWord.ActiveDocument.Sections(A_Index).Footers(1).Range.Delete
			counter++
	}

splashNotify("Removed " . counter . " items in " . dName) 
Sleep 2000
ExitApp 

ChangeButtonNames: 
IfWinNotExist, Add or Delete
    return  ; Keep waiting.
SetTimer, ChangeButtonNames, off 
WinActivate 
ControlSetText, Button1, &Headers
ControlSetText, Button2, &Footers 
ControlSetText, Button3, &Both
return
