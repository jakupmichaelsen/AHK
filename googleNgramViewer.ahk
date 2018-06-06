#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
SendMode Input
TrayTip, %A_ScriptName%, Script loaded
#NoTrayIcon

splashText("Google Books Nram Viewer", , "Enter comma-separated search items",  "Albert Einstein,Sherlock Holmes,Frankenstein")
if input <>
    run https://books.google.com/ngrams/graph?content=%input%&year_start=1800&year_end=2000
Return