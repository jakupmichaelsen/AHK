/*
This script allows you to merge diagrams, charts, and any other information saved as bitmaps into a range of word documents.
The reason for working with bmp, instead of png or jpg, is because of MS Office's blurring of exported images in other formats than bitmaps. 
Insert a bookmark using Word's 'Insert Bookmark' option wherever you  would like the image to appear. Name each bookmark as "b1", "b2", "b3", and so on.
Save your bitmaps in accordingly-numbered subfolders "\1", "\2", "\3", etc.; and your Word documents in a "\doc" folder. 
The script then asks you for the root folder containing all of your data, and created to new subfolders - "\merged" and "\pdf" - each with a pdf and Word format file of your merged documents. 
*/

global oWord := ComObjCreate("Word.Application") 


FileSelectFolder, Folder, , , Select the root folder containing your Word documents and bitmaps subfolders
getBitmapFolders(Folder)
FileCreateDir, %Folder%\pdf
FileCreateDir, %Folder%\merged
mergeDocs(Folder)
oWord.quit
Msgbox Done!
Run, %Folder%
ExitApp, 

mergeDocs(Folder)
{
	global
	loop, %Folder%\doc\*.*
 	{
 		docIndex = %A_Index%
 		; MsgBox, docIndex: %docIndex%
 		StringGetPos, pos, A_LoopFileName, ~
 		if pos = 0
  			Continue
  		Else
  		{
	 		oWord.Application.Documents.Open(A_LoopFileFullPath)
	 		; MsgBox, number of bitmap folders:  %bitmaps%
	 		loop, %bitmaps%
	 		{
		 		bmpIndex = %A_Index%
		 		; MsgBox, bmpIndex: %bmpIndex%
		 		oWord.Application.Selection.Find.ClearFormatting()
		 		bookmark = b%bmpIndex%
		 		; MsgBox, %bookmark%
				oWord.Application.ActiveDocument.Bookmarks(bookmark).Select
		 		FileName = %Folder%\%bmpIndex%\%docIndex%.bmp
		 		oWord.Selection.InlineShapes.AddPicture(FileName)
		 	}
	 		SplitPath, A_LoopFileFullPath, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive
		 	oWord.Application.ActiveDocument.SaveAs2(Folder . "\pdf\" . OutNameNoExt . ".pdf", 17) 
			oWord.Application.ActiveDocument.SaveAs(Folder . "\merged\" . A_LoopFileName)
			oWord.Application.ActiveDocument.Close()
		}
 	}
}

getBitmapFolders(Folder)
{
	global bitmaps = 0
	Loop, %Folder%\*.*,2
		{
		if A_LoopFileName is integer
			bitmaps++
	    Else
	    	Continue
		}
}
