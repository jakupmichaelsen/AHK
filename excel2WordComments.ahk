try
    xl := ComObjActive("Excel.Application") 
catch
{
    MsgBox,16, %A_ScriptName%, No active Excel document found. `rOpen your comments sheet and try again.
    ErrorLevel = 1
}
if ErrorLevel <> 1 ; Run only if ErrorLevel isn't = 1
{
    refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z") ; array of column letters
     
    titleColumn(xl) ; Display title column and return %choice% 
     
    if choice <> ; Run only if choice isn't empty
    {
        Loop, % xl.ActiveSheet.UsedRange.Rows.Count  { ; Loop through used rows
            if xl.Range(refColumns[1] . A_Index).Value = choice
                row := A_Index
        }
        Loop, % xl.ActiveSheet.UsedRange.Columns.Count { ; Loop through used columns 
            cell := xl.Range(refColumns[A_Index] . row).Value 
            if A_Index > 1 ; Skip header-row
            {
                if cell <>
                {
                    if A_Index = 2
                        text := cell 
                    if A_Index = 3
                        header := cell
                    if A_Index = 4
                        URL := cell
                    if A_Index = 5
                        title := cell
                }
            }
        }
    }
 
   if choice <> ; Run only if choice isn't empty
       wComment(text, header, URL, title) ; Add comment 

}
ExitApp, 
