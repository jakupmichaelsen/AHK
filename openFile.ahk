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
	Run, %comspec% /c "cd %dir% && reveal-md %file%"
Else 
	run %dir%\%file%