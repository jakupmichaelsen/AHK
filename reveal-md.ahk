Loop %0%  ; For each file dropped onto the script (or passed as a parameter).
{
    GivenPath := %A_Index%  ; Retrieve the next command line parameter.
    Loop %GivenPath%, 1
    {
        file := A_LoopFileName
    	MsgBox, %file%
    	dir := A_LoopFileDir
    	MsgBox, %dir%
    }
	Run, %comspec% /k "cd %dir% && reveal-md %file%"
}
