#SingleInstance, force
oWord := ComObjCreate("Word.Application")
xl := ComObjActive("Excel.Application") 


FileSelectFolder, Folder, , , Select the folder containing your Word documents
FileCopyDir, %Folder%, %Folder%\merged, 1
Folder = %Folder%\merged

Loop, % xl.Worksheets.Count
	{
		xl.Worksheets(A_Index).Activate
		rowRange := xl.Worksheets(A_Index).Name . "Row"
		filenameRange := xl.Worksheets(A_Index).Name . "Filename"
		bookmark := xl.Worksheets(A_Index).Name
		Loop, % xl.Worksheets(A_Index).UsedRange.Rows.Count - 1 ; Loop through data-rows (= used rows minus the header-row)
		{
			xl.Range(rowRange).Value := A_Index
			xl.ActiveSheet.ChartObjects(1).Copy
			fName := % folder . "\" . xl.Range(filenameRange).Value . ".docx"
			oWord.Application.Documents.Open(fName)
			oWord.Application.ActiveDocument.Bookmarks(bookmark).Select
			oWord.Selection.Paste 
			oWord.Application.ActiveDocument.Save()
			oWord.Application.ActiveDocument.Close() 
		}
}

MsgBox, Done!
oWord.quit