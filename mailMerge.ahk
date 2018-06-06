
; Word mail merge - from active Excel sheet to individual docx files. 
; Just add bookmarks in your Word template where you want the data to appear, and name them according to the Excel data headers.
; Using bookmarks avoids the 256 character limit of Office' search-and-replace function 

#SingleInstance, force
xl := ComObjCreate("Excel.Application") 
oWord := ComObjCreate("Word.Application")
refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z") ; array of column letters

docMergeGUI()
WinWaitClose, ahk_class AutoHotkeyGUI 

xl.Application.Workbooks.Open(data)
filenameColumn(xl, refColumns) ; Display column headers and return %choice% 
selectRows(xl)

Loop, % xl.ActiveSheet.UsedRange.Rows.Count ; Loop through used rows
	{
		if selectRow%A_Index% = 1 ; Merge only selected rows
		{
			fileName := xl.Range(refColumns[choice] . A_Index).Value ; Get value from filename column
			mergeFile = %outDir%\%fileName%.docx
			TrayTip, mailMerge.ahk, Copying template to %mergeFile%
			FileCopy, %template%, %mergeFile% , 1 ; Copy template to new file, overwriting any existing 'mergefile'
			oWord.Application.Documents.Open(mergeFile)
			mergeRow(A_Index) ; <- This is the actual merge function  
			TrayTip, mailMerge.ahk, Saving
			oWord.Application.ActiveDocument.Save()
			oWord.Application.ActiveDocument.Close()
		}
	}

MsgBox, Done!
xl.quit
oWord.quit

mergeRow(row){
	global 
	nBookmarks := oWord.Application.ActiveDocument.Bookmarks.Count ; Count number of bookmarks
	loop, %nBookmarks% ; Create separate variables for each bookmark 
		bookmark%A_Index% := oWord.Application.ActiveDocument.Bookmarks(A_Index).Name
		
	loop, %nBookmarks% ; Loop through bookmarks 
	{
		bookmark = % bookmark%A_Index% ; bookmark = bookmark variables 
		Loop, % xl.ActiveSheet.UsedRange.Columns.Count ; Loop through used columns 
		{
			header := xl.Range(refColumns[A_Index] . 1).Value ; Get headers 
			if header = %bookmark% ; If a header matches a bookmark name: 
			{
				clipboard := ; Clear clipboard
	 			clipboard := xl.Range(refColumns[A_Index] . row).Value ; Copy cell value
	 			if clipboard <> ; If cell, and hence clipboard, not empty:
				{
					oWord.Application.ActiveDocument.Bookmarks(bookmark).Select ; Select bookmark in Word
					if clipboard is Number
						clipboard := Floor(clipboard)
					oWord.Selection.Paste ; Paste cell value, overwriting bookmark	
				}
			}
		}
	
	}
} 	

