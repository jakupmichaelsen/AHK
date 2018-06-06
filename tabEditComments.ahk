#SingleInstance, force 
#NoTrayIcon
	
oWord := ComObjActive("Word.Application")

count = 1

#IfWinActive, ahk_class OpusApp
^Tab:: 
	nComments := oWord.ActiveDocument.Comments.Count
	oWord.ActiveDocument.Comments(count).Edit
	count++
	if count > %nComments%
		Reload
	Return
+^Tab:: 
	nComments := oWord.ActiveDocument.Comments.Count
	count--
	if count < 1
		count = %nComments%
	oWord.ActiveDocument.Comments(count).Edit
	Return



