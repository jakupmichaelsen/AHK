#SingleInstance, force
AutoReload()
TrayTip, %A_ScriptName%, Loaded

SetTitleMatchMode, 2

#If WinActive("Google Docs - Google Chrome")
+F2::Click 3 ; Select paragraph

