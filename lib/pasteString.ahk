pasteString(string){
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	Clipboard := string 
	Send ^v 
	Sleep 50
	Clipboard = %ClipSaved%  ; Restore clipboard
}