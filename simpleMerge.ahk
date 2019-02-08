refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z") ; array of column letters

xl := ComObjActive("Excel.Application")
oWord := ComObjActive("Word.Application")

Loop, % xl.ActiveSheet.UsedRange.Rows.Count ; Loop through used rows
{
	if A_Index > 1 ; Skip header row 
	{
		fileName := xl.Range("A" . A_Index).Value ; Get value from ilename column
		msgbox %fileName%
		mergeFile = d:\%fileName%.docx
		msgbox %mergeFile%
		TrayTip, mailMerge.ahk, Copying template to %mergeFile%
		FileCopy, %template%, %mergeFile% , 1 ; Copy template to new file, overwriting any existing 'mergefile'
		oWord.Application.Documents.Open(mergeFile)
		mergeRow(A_Index) ; <- This is the actual merge function  
		TrayTip, mailMerge.ahk, Saving
		oWord.Application.ActiveDocument.Save()
		oWord.Application.ActiveDocument.Close()
	}
}

