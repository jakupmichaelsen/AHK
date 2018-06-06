#SingleInstance, force
F1::
    xl := ComObjActive("Excel.Application") 
    Loop, 18 ; % xl.ActiveSheet.UsedRange.Rows.Count - 1 ; Loop through data-rows (= used rows minus the header-row)
    {
        xl.ActiveSheet.Range("rubricRow").Value := A_Index
        fName := xl.ActiveSheet.Range("rubricFilename").Value 
       	fPath = D:\Dropbox\Code\ahk\Word\%fName%.bmp
       	xl.ActiveSheet.ChartObjects(1).Chart.Export(fPath)
    }
    Return