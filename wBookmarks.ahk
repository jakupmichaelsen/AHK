; Inspired from the way of handling bookmarks in the texteditor PSPad I wrote the following script to adapt this to Microsoft Word.
; 
; The following hotkeys are defined:
; Set a bookmark <Alt> + <Left-arrow key>
; Jump up to next bookmark <Alt> + <Up-arrow key>
; Jump down to next bookmark <Alt> + <Down-arrow key>
; Delete a bookmark <Alt> + <Right-arrow key> 
; 
; If you want to delete a bookmark, you have to jump first with <Alt> + <Up> or <Alt> + <Down> to the bookmark and then press <Alt> + <Right> to delete it.
; 
; The script sets new Word bookmarks in the following way: b1, b2, b3, etc.
; 
; Tested with MS Word 2010 and WinXP/Win7, requires AHK_L
; 
; Issues: One bookmark could be defined multiple
; 
; Perhaps you have some suggestions to improve this script. :D 


#SingleInstance force  
#NoTrayIcon
AutoReload()

#IfWinActive, ahk_class OpusApp

F2::Gosub, nextBookmark
^F2::Gosub, setBookmark
+^F2::Gosub, delBookmark

nextbookmark:
	oWord := ComObjActive("Word.Application") 
	bookmarks_Count := oWord.ActiveDocument.Bookmarks.count ; count number of defined bookmarks
	If (bookmarks_Count > 0) {
		If (bookmark_i < bookmarks_Count AND bookmark_i > 0) {
	 		bookmark_i++
			oWord.ActiveDocument.Bookmarks(Index := bookmark_i).select  ; jump to i.-bookmark
		}
		Else If (bookmark_i = bookmarks_Count) { 
			oWord.ActiveDocument.Bookmarks(Index := 1).select  ; jump to first bookmark
			bookmark_i := 1
	  }
	  Else { ; if bookmark_i = 0 or undefined 
	    oWord.ActiveDocument.Bookmarks(Index := bookmarks_Count).select  ; jump to the last bookmark
	  	bookmark_i := bookmarks_Count 
	  }
	  splashNotify(%bookmark_i% "bookmark", 1500, "bottom")
	}
	Return


setBookmark:
	Loop { ; looking for a free bookmarkname
		If oWord.ActiveDocument.Bookmarks.Exists("b" A_Index) 
			continue
		temp := "b" A_Index		
		break
	}
	oWord.ActiveDocument.Bookmarks.Add( temp, oWord.Selection.Range )
	bookmark_i := oWord.ActiveDocument.Bookmarks(temp).Range.BookmarkID
	splashNotify(%bookmark_i% "bookmark added", 1500, "bottom")
	Return 

delBookmark:
	If (bookmark_i > 0)	{
		oWord := ComObjActive("Word.Application") 
		Try 
			act_bName := oWord.ActiveDocument.Bookmarks(bookmark_i).Name
		If oWord.ActiveDocument.Bookmarks.Exists(act_bName) 
		{
			oWord.ActiveDocument.Bookmarks(act_bName).Delete
			splashNotify(%bookmark_i% "bookmark deleted", 1500, "bottom")
		}
	}
	Return

+!Up::  ; jump up to next bookmark
	oWord := ComObjActive("Word.Application")
	bookmarks_Count := oWord.ActiveDocument.Bookmarks.count ; count number of defined bookmarks
	If (bookmarks_Count > 0) {
		if (bookmark_i > 1) {
			bookmark_i--
			oWord.ActiveDocument.Bookmarks(Index := bookmark_i).select   ; jump to i.-bookmark
		}
	  Else If (bookmark_i = 1) { 
	    oWord.ActiveDocument.Bookmarks(Index := bookmarks_Count).select  ; jump to the last bookmark
			bookmark_i := bookmarks_Count  
	  }
	  Else { ; if bookmark_i = 0 or undefined   
	    oWord.ActiveDocument.Bookmarks(Index := 1).select  ; jump to first bookmark
	    bookmark_i := 1
	  }
	  splashNotify(%bookmark_i% "bookmark", 1500, "bottom")
	}
	Return