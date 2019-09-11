#SingleInstance, force 

oWord := ComObjActive("Word.Application")
nComments := oWord.ActiveDocument.Comments.Count
cCount = 1
#IfWinActive, ahk_class OpusApp
Tab:: 
	oWord.ActiveDocument.Comments(cCount).Edit
	cCount++
	if cCount > %nComments%
		Reload
	Return
+Tab:: 
	cCount--
	if cCount < 1
		cCount = %nComments%
	oWord.ActiveDocument.Comments(cCount).Edit
	Return



