#NoEnv ; Significantly improves performance by not looking up empty variables as environmental variables 
#InstallMouseHook
#InstallKeybdHook
#SingleInstance, force
#NoTrayIcon
AutoReload()

splashNotify(A_ScriptName " reloaded") 
IniRead, browser, splashUI.ini, paths, browser
IniRead, onedrive, splashUI.ini, paths, onedrive
IniRead, user, splashUI.ini, paths, user
IniRead, ahk, splashUI.ini, paths, ahk
IniRead, git, splashUI.ini, paths, git


/*
  _________ _______________________________
 /   _____//   _____/\__    ___/\__    ___/
 \_____  \ \_____  \   |    |     |    |   
 /        \/        \  |    |     |    |   
/_______  /_______  /  |____|     |____|   
        \/        \/                       
*/

GroupAdd, ssttIgnore, ahk_class PX_WINDOW_CLASS
GroupAdd, ssttIgnore, ahk_exe PDFXEdit.exe
GroupAdd, ssttIgnore, ahk_class XLMAIN

#IfWinActive, ahk_class XLMAIN 
+^k::
	xl := ComObjActive("Excel.Application")
	xl.Selection.Delete(Shift := -4162) ; xlUp = -4162 
	Return 

+!Left::
	xl := ComObjActive("Excel.Application")
	xl.Selection.Delete(Shift := -4159) ; xlToLeft = -4159 
    ; Selection.Insert Shift:=xlToRight ; xlToRight = -4161
	Return 

#IfWinActive

#IfWinNotActive, ahk_group ssttIgnore
+^Delete::
	Send, +{End}{Delete}
	Return
+^BackSpace:: ; Delete whole line
	Send, +{Home}{Delete}
	Return
+^d::
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Sleep, 100 
	Send, {Home}
	Sleep, 100 
	Send, +{End}
	Sleep, 100 
	Send, ^c
	Send, {End}
	Sleep, 100 
	Send, {Enter}
	Send, ^v
	Sleep, 100
	Clipboard = %ClipSaved%  ; Restoreore clipboard
	Return
+^k::
	Send, {Home}
	Send, +{End}{Delete}
	; Send, {Delete}
	Return

