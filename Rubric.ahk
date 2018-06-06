#SingleInstance, force 
Menu, Tray, Icon, %A_ScriptDir%\icons\R.ico

xl := ComObjActive("Excel.Application") 
refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z")

Gui, -Caption -Border +AlwaysOnTop +OwnDialogs
Gui, Color, 221A0F, 221A0F
Gui, Margin, 10, 5
Gui, Font, s13 bold cFF9900, Courier New
Gui, Add, Text, Center , %A_ScriptName%
Gui, Font, s2 cFF9900, Consolas
Gui, Add, Text, 
Gui, Font, s12 normal cFF9900, Consolas
Gui, Add, Edit, vClass -E0x200 -WantReturn  Limit, [class] ; Press ctrl+enter for new line and enter to submit
Gui, Add, Edit, x+20 vStudent -E0x200 -WantReturn , [student] ; Press ctrl+enter for new line and enter to submit
Loop, % xl.ActiveSheet.UsedRange.Columns.Count ; Loop through used columns
{
	header := xl.Range(refColumns[A_Index] . 1).Value
		if header not in Timestamp,Class,Student,Grade ; Comma-separated list of columns to skip ("space sensitive")
		{
			IfInString, header, & ; If the header contains an ampersand
				StringReplace, header, header, &, && ; Replace with two &s - needed to display it correctly 
			Gui, Add, Text, x20, %header%
			Gui, Add, Slider, vMySlider%A_Index% +NoTicks Range1-4 Reverse Thick15 TickInterval ToolTip, 1
		}
}
Gui, Font, s4 cFF9900, Consolas
Gui, Add, Text, 
Gui, Font, s12 cFF9900, Consolas
Gui, Add, Text, x40 gOK Center, Submit 
Gui, Add, Text, x+20 gExit Right, Exit 
Gui, Add, Button, x-100 y-100, OK 
          ; OK button is hidden off the canvas, but is needed for submitting with enter.

Gui, Show, x0 y0 H768

WinWaitClose, ahk_class AutoHotkeyGUI 
GuiEscape:
GuiClose:
Exit:
Gui, Destroy
ExitApp

ButtonOK:
OK:
Gui, Submit  ; Save each control's contents to its associated variable.
Gui, Destroy

newRow := 1 + xl.ActiveSheet.UsedRange.Rows.Count ; New row number 
Loop, % xl.ActiveSheet.UsedRange.Columns.Count ; Loop through used columns
{
	if xl.Range(refColumns[A_Index] . 1).Value = "Timestamp" 
	{
		FormatTime, TimeString, YYYYMMDDHH24MISS, dd.MM.yy   HH:mm ; Enter timestamp 
		xl.Range(refColumns[A_Index] . newRow).Value := TimeString
	}
	else if xl.Range(refColumns[A_Index] . 1).Value = "Student"
			xl.Range(refColumns[A_Index] . newRow).Value := Student
	else if xl.Range(refColumns[A_Index] . 1).Value = "Class"
			xl.Range(refColumns[A_Index] . newRow).Value := Class
	Else
	{
		score := MySlider%A_Index% ; Get score from respective slider 
		xl.Range(refColumns[A_Index] . newRow).Value := score
	}
}
xl.ActiveWorkbook.Save  ; Save Workbook and goto end of data 
xl.Selection.End(Unit:=1).Select 
xl.Selection.End(Unit:=-4121).Select

Reload ; Reload script until user selects Exit 