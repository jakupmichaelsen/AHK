#SingleInstance, force
refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z") 
Progress, CT%header_color% CW%background% B1 Y0 ZH0 CTFF0000, Navigate to relevant folder and press F1 to copy filenames to 'Filenames.xlsx' (ESC to cancel)
xl := ComObjCreate("Excel.Application") 
xl.Application.Workbooks.Add() 
column = 0

GetNames:
column++

Esc::
	xl.Quit()
	xl := 
	ExitApp
F1::
folder := Explorer_GetPath()
Loop, %folder%\*.*
{
	SplitPath, A_LoopFileFullPath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
	xl.Range(refColumns[column] . A_Index).Value := OutFileName
}
path = %folder%\Filenames.xlsx
xl.ActiveWorkbook.SaveAs(path)
Return 