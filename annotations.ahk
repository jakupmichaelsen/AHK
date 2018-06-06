#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
SendMode Input
TrayTip, %A_ScriptName%, Script loaded
AutoReload()
xl := ComObjActive("Excel.Application") 
refColumns := Object(1,"A",2,"B",3,"C",4,"D",5,"E",6,"F",7,"G",8,"H",9,"I",10,"J",11,"K",12,"L",13,"M",14,"N",15,"O",16,"P",17,"Q",18,"R",19,"S",20,"T",21,"U",22,"V",23,"W",24,"X",25,"Y",26,"Z")  

#IfWinActive, ahk_class PXE:{C5309AD3-73E4-4707-B1E1-2940D8AF3B9D} ; ahk_class AcrobatSDIWindow
MButton::
F2::
    ClipSaved = %ClipboardAll%  ; Save clipboard
    Clipboard = ; Clear clipboard 
    send ^c 
    ClipWait, 0.2, 1
    WinGetActiveTitle, doc 
    sheets = ; Clear sheets
    choice = ; Clear choice

try
    row := xl.ActiveSheet.UsedRange.Rows.Count + 1
catch
{
    MsgBox,16, %A_ScriptName%, Could not loop through active Excel document. `rMake sure you're not editing a cell and try again.
    ErrorLevel = 1
}
if ErrorLevel <> 1 ; Run only if ErrorLevel isn't = 1
{
    tags =
    quote := Clipboard
    Clipboard = %ClipSaved%  ; Restore clipboard

    Loop, % xl.ActiveSheet.UsedRange.Rows.Count { ; Loop through used columns 
            if A_Index > 1 ; Skip header-row
            {
                docCell := xl.Range("A" . A_Index).Value
                tagCell := xl.Range("C" . A_Index).Value
                if docCell = %doc%
                {
                IfNotInString, tags, %tagCell%   
                    tags .= xl.Range("C" . A_Index).Value "|" ; Get column headers
                }
            }
    }
    ;  StringTrimRight, tags, tags, 1
    cleanClipboard()
    splashMenu("annotations.ahk", tags, quote)
    if (choice <> "" or input <> "")
    {
        ; Click, right 
        ; Send {AppsKey}
        Send !h
        ;  Send v
        xl.ActiveSheet.Range("A" . row).Value := doc
        xl.ActiveSheet.Range("B" . row).Value := quote
        if input <>
            xl.ActiveSheet.Range("C" . row).Value := input
        Else
            xl.ActiveSheet.Range("C" . row).Value := choice
        ; xl.ActiveSheet.Range("D" . row).Value := notes
        xl.ActiveSheet.Range("E" . row).Value := A_Now
    }
}
Return 

splashMenu(text, list, helptext = ""){ ; Pass help text and |-separated list
  global
  choice =  ; Clear previous choices
  rows = ; Clear previous number of rows
  input = ; Clear previous input
  Loop, Parse, list, | ; Calculate ListBox height 
    rows = %A_Index% 
  if rows < 1 ; 1 row if no tag-list
    rows = 1
  boxWidth := StrLen(text) * 25 ; and width 
  IniRead()
  Gui, splashMenu: -Caption -Border +AlwaysOnTop +OwnDialogs
  Gui, splashMenu: Color, %backC%, %backC%
  Gui, splashMenu: Margin, 10, 10
  Gui, splashMenu: Font, s20 bold c%secondaryC%, %font%
  Gui, splashMenu: Add, Text, , %text%
  Gui, splashMenu: Font, italic s8 c%secondaryC%
  Gui, splashMenu: Add, Text, w%boxWidth%, Clipboard:`r %helptext% 
  Gui, splashMenu: Font, s%size% %style% c%fontC%
  Gui, splashMenu: Add, Text, , Tags
  Gui, splashMenu: Add, Edit, vinput -WantReturn -VScroll w%boxWidth%
  Gui, splashMenu: Add, ListBox, ReadOnly vchoice r%rows% gTagBox w%boxWidth%, %list% 
  ; Gui, splashMenu: Add, Edit, r5 vnotes -WantReturn -VScroll w%boxWidth%
  Gui, splashMenu: Add, Button, x-100 y-100 Default, OK ; Hide default button, as we're going to use enter to submit
  Gui, splashMenu: Show, x0 y0 h%A_ScreenHeight%, GUI splashMenu ; Create and display the GUI  
  WinWaitClose, ahk_class AutoHotkeyGUI 
  Return 

  
  TagBox:
  if A_GuiControlEvent <> DoubleClick
    return
  ; Otherwise, the user double-clicked a list item.
  Gui, splashMenu: Submit
  Gui, splashMenu: Destroy
  Return
  
  splashMenuGuiEscape:
  choice = 
  input =  
  Gui, splashMenu: Destroy
  Return 

  splashMenuButtonOK:
  Gui, splashMenu: Submit
  Gui, splashMenu: Destroy
  Return
}