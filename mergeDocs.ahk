#SingleInstance, force
Progress, CT%header_color% CW%background% B1 Y0 ZH0 CTFF0000, Navigate to relevant folder and press F1 to merge doc(x)s (ESC to cancel)


Esc::
	xl.Quit()
	xl := 
	ExitApp
F1::
folder := Explorer_GetPath()
path = %folder%\Filenames.xlsx
Extensions := "doc,docx"

Folder := RTrim(Folder,"")
oWord := ComObjCreate("Word.Application")
oDoc := oWord.Documents.Add

Loop, %Folder%\*.*,0,1
{
   if A_LoopFileExt in %Extensions%
   {
      if (A_index != 1)
         oDoc.Content.InsertParagraphAfter
      oDoc.Content.InsertAfter(A_LoopFileFullPath)
      oDoc.Content.InsertParagraphAfter
      oRange := oDoc.Content
      oRange.Collapse(0)   ; wdCollapseEnd = 0
      oRange.InsertFile(A_LoopFileFullPath)
   }   
}
oWord.Visible := 1, oWord.Activate
MsgBox,,, Done!,1
ExitApp
