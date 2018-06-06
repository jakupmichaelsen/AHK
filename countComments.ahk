#SingleInstance, force 

oWord := ComObjActive("Word.Application")
nComments := oWord.ActiveDocument.Comments.Count
count = 1
#IfWinActive, ahk_class OpusApp
Tab:: 
	oWord.ActiveDocument.Comments(count).Edit
	count++
	if count > %nComments%
		Reload
	Return
+Tab:: 
	count--
	if count < 1
		count = %nComments%
	oWord.ActiveDocument.Comments(count).Edit
	Return



