#SingleInstance, force
AutoReload()
Menu, Tray, Icon, D:\Dropbox\Code\AHK\icons\C.ico

; Styles we want to remove from the console window:
WS_POPUP := 0x80000000
WS_CAPTION := 0xC00000
WS_THICKFRAME := 0x40000
WS_EX_CLIENTEDGE := 0x200

; Styles we want to add to the console window:
WS_CHILD := 0x40000000

; Styles we want to add to the Gui:
WS_CLIPCHILDREN := 0x2000000

; Flags for SetWindowPos:
SWP_NOACTIVATE := 0x10
SWP_SHOWWINDOW := 0x40
SWP_NOSENDCHANGING := 0x400

; Create Gui and get window ID.
Gui, +LastFound +%WS_CLIPCHILDREN%
GuiWindow := WinExist()
Gui, -Border +OwnDialogs 
Gui, Color, 221A0F, 221A0F
Gui, Margin, 50, 25
Gui, Font, s12 bold cFF9900, Consolas

; Launch hidden cmd.exe and store process ID in pid.
Run, %ComSpec%,, Hide, pid

; Wait for console window to be created, store its ID.
DetectHiddenWindows, On
WinWait, ahk_pid %pid%
ConsoleWindow := WinExist()

; Get size of console window, excluding caption and borders:
VarSetCapacity(ConsoleRect, 16)
DllCall("GetClientRect", "uint", ConsoleWindow, "uint", &ConsoleRect)
ConsoleWidth := NumGet(ConsoleRect, 8)
ConsoleHeight:= NumGet(ConsoleRect, 12)

; Apply necessary style changes.
WinSet, Style, % -(WS_POPUP|WS_CAPTION|WS_THICKFRAME)
WinSet, Style, +%WS_CHILD%
WinSet, ExStyle, -%WS_EX_CLIENTEDGE%

; Put the console into the Gui.
DllCall("SetParent", "uint", ConsoleWindow, "uint", GuiWindow)

; Move and resize console window. Note that if SWP_NOSENDCHANGING
; is omitted, it incorrectly readjusts the size of its client area.
DllCall("SetWindowPos", "uint", ConsoleWindow, "uint", 0
    , "int", 10, "int", 10, "int", ConsoleWidth, "int", ConsoleHeight
    , "uint", SWP_NOACTIVATE|SWP_SHOWWINDOW|SWP_NOSENDCHANGING)

; Show the Gui. Specify width since auto-sizing won't account for the console.
Gui, Show, x0 y30 h768 w1366, AHK CMD

; Be sure to close cmd.exe later.
; OnExit, Exiting

; If cmd.exe exits prematurely, fall through to ExitApp below.
Process, WaitClose, %pid%

GuiEsape:
GuiClose:
; ButtonOK:
; Exiting:
OnExit
Process, Close, %pid% ; May be a bit forceful? No effect if it already closed.
ExitApp