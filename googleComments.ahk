; ============= AUTOEXEC SECTION  =============
#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
AutoReload()
Menu, Tray, Icon, D:\Dropbox\Code\AHK\icons\g.ico

SendMode, InputThenPlay
try
    xl := ComObjActive("Excel.Application") 
catch
    run %A_ScriptDir%\myComments.xlsx

xl := ComObjActive("Excel.Application") 
refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z") ; array of column letters
global toggle := 0 ; Global variable used for toggling 

SetTitleMatchMode, 2
#IfWinActive, Google Docs - Google Chrome

F3:: ; Brings up a list of standard comments and inserts them 
    Top:
    choice =
    categorySheets(xl) ; Show selection of worksheets and choice which to activate 
    if choice <> ; Run only if choice isn't empty
    {
        xl.Worksheets(choice).Activate
        titleColumn(xl) ; Display title column and return %choice% 
        
        if choice <> ; Run only if choice isn't empty
        {
            Loop, % xl.ActiveSheet.UsedRange.Rows.Count  { ; Loop through used rows
                if xl.Range(refColumns[1] . A_Index).Value = choice
                    row := A_Index
            }
            Loop, % xl.ActiveSheet.UsedRange.Columns.Count 
            { ; Loop through used columns 
                xl.Range(refColumns[A_Index] . row).Select
                cell := xl.Selection.Value
                
                if A_Index > 1 ; Skip header-row
                {
                    if cell <>
                    {
                        if A_Index = 2
                            comment := cell
                        if A_Index = 3
                            header := cell
                        if A_Index = 4
                            URL := cell
                    }
                }
            }
            Send %comment%`r`r%URL%
        }
        Else
            Gosub, Top
        
        header = 
        text = 
        URL =
        title = 
        row = 
    }
Return



