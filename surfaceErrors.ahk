#SingleInstance, force 
IniRead, ahk, splashUI.ini, paths, ahk

refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z") 

errors = %ahk%\errors.txt 
txtRadioOptions(errors)

try
	xl := ComObjActive("Excel.Application") 
catch
{
	MsgBox,16, %A_ScriptName%, No active Excel document found.`rExiting.
	ExitApp
}



CapsLock::
	splashUI("r2", "Alt + N: insert new row", options)
	if choice <> 	
	{
		Loop, % xl.Sheets(1).UsedRange.Columns.Count ; Loop through used columns
		{
			if xl.Sheets(1).Range(refColumns[A_Index] . 1).Value = choice 
				xl.Sheets(1).Range(refColumns[A_Index] . 2).Value += 1
		}
		oWord := ComObjActive("Word.Application") 
		oWord.Application.ActiveDocument.Bookmarks.Add(Name:="bookmark")
		oWord.Selection.Expand(Unit:=3)
		oWord.Selection.Copy
		FormatTime, CurrentDateTime,, yyyyMMddHHmm
		fileName := oWord.ActiveDocument.Name
		FileAppend, %CurrentDateTime%`;%fileName%`;%choice%`;%clipboard%`r, fejlsætninger.txt
		; WinActivate, ahk_class OpusApp
		oWord.Application.ActiveDocument.Bookmarks("bookmark").Select
		oWord.ActiveDocument.Bookmarks("bookmark").Delete
		Return
	}
	Return

#IfWinActive, ahk_class AutoHotkeyGUI
!n::
	Send, {Esc}
	WinActivate, ahk_class XLMAIN
	xl.Sheets(1).Rows(2).Insert ; Sheets(1)
	Loop, % xl.Sheets(1).UsedRange.Columns.Count ; Loop through used columns
	{
		xl.Sheets(1).Range(refColumns[A_Index] . 2).Value := 0
	}
	Return