#SingleInstance, force
Progress, CT%header_color% CW%background% B1 Y0 ZH0 CTFF0000, Navigate to relevant folder and press F1 to rename filenames from column A to column B in 'Filenames.xlsx' (ESC to cancel)


refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z") 
xl := ComObjCreate("Excel.Application") 

Esc::
	xl.Quit()
	xl := 
	ExitApp
F1::
folder := Explorer_GetPath()
path = %folder%\Filenames.xlsx
xl.Application.Workbooks.Open(path)

Loop, % xl.ActiveSheet.UsedRange.Rows.Count ; Loop through used rows
	{
		if A_Index > 1 ; Skip header row 
		{
			oldFilename := folder . "\" . xl.Range("A" . A_Index).Value 
			newFilename := folder . "\" . xl.Range("B" . A_Index).Value 
			FileMove, %oldFilename%, %newFilename%, 1
		}
	}
xl.Quit()
xl := 
ExitApp
