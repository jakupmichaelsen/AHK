F1::
	splashText("Paste your text here", 4) ; Popup asking to paste text
	msgbox %input%
	sansSpecialChar := RegExReplace(input, "[^a-zA-Z0-9\s*\.,?:!;]", "") ; Remove all save the characters in []
	Loop, parse, sansSpecialChar, `.`::`;`,?!, %A_Space% ; Parse text and break on .:;,? and !
	{
	        FileAppend, %A_LoopField%`n, D:\test1.txt
	}
	Return

F2::
	Loop, read, c:\wordlist.txt 
		{
			word = %A_LoopReadLine%
			Loop, read, D:\test1.txt, D:\test2.txt
			{
	    		IfInString, A_LoopReadLine, %word%, FileAppend, %A_LoopReadLine%`n
			}
		} 
	Return