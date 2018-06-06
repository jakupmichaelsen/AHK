explorerPath() ; Get active Windows Explorer window's path and store it in %path% variable 
{
	global
	WinGetText, txt, A, 
	Loop, Parse, txt, `r
	{
		if A_Index = 1
			path := A_LoopField
	}
	StringTrimLeft, path, path, 9
}

