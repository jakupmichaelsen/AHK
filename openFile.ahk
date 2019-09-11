FileRead, rootDir, lastDir.txt
FileSelectFile, file,, %rootDir%
Loop, %file%
{
	file := A_LoopFileName
    dir := A_LoopFileDir
	extension := A_LoopFileExt
}
FileDelete, lastDir.txt
FileAppend, %dir%, lastDir.txt
if extension = md
{
	RunAs, laerer, 2512
	Run, %comspec% /k "cd %dir% && reveal-md %file%"
	; Send, #r
	; WinWaitActive, ahk_class #32770
	; Send, cmd {Enter} 
	; Sleep 1000 
	; Send, cd %dir% {Enter} 
	; Send, reveal-md %file% {Enter}
}
Else 
	run %dir%\%file%