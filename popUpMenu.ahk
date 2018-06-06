#Include, %A_ScriptDir%\menus
; Include sub menus 
#Include, bloomsVerbs.ahk
#Include, warmers.ahk

; Main popup menu
Menu, TopMenu, Add, Bloom's Verbs, :Bloom's
Menu, TopMenu, Add, Warmers, :Warmers 
Menu, TopMenu, Add, 
Menu, TopMenu, Add, Copy Filenames, copyFilename
Menu, TopMenu, Add, Kongruensfejl, Kongruensfejl
Menu, TopMenu, Add, Grading Papers, gradingpapers
gradingpapers:
    Run, GradingPapers.ahk
    Return


; Include menu actions 
#Include, bloomsVerbsActions.ahk
#Include, warmersActions.ahk
#Include, commandsActions.ahk
Return
#F2::Menu, TopMenu, Show, 800, 300

