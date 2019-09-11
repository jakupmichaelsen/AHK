splashInput(text, rows = 1, helptext = "", defaultTxt = "", timeout = ""){
	global  ; Set variable as global - rename 'vinput' below if you use 'input' as a variable elsewhere.
	input = ; Clear input variable
	IniRead()
	guiW := A_ScreenWidth / 3
	guiX := A_ScreenWidth - guiW 
	Gui, splashInput: -Caption -Border +AlwaysOnTop +OwnDialogs
	Gui, splashInput: Color, %background%, %background%
	Gui, splashInput: Margin, 50, 25 
	Gui, splashInput: Font, s%size% %style% c%header_color%, %typingFont%
	Gui, splashInput: Add, Edit, r%rows% w%guiW% vinput -WantReturn -E0x200 -VScroll, %defaultTxt%
		  ; Edit box with variable rows. Press ctrl+enter for new line and enter to submit. 
	Gui, splashInput: Font, s10 c%font_color%, %header%
	if helptext <> ; If helptext not empty 
		Gui, splashInput: Add, Text,, %helptext%
	Gui, splashInput: Add, Button, x-100 y-100 Default, OK 
		  ; OK button is hidden off the canvas, but is needed for submitting with enter.
	Gui, splashInput: +OwnDialogs
	Gui, splashInput: Show, x%guiX% w%guiW%  
	splashNotify(text, "middle", 600000, 22, 200) ; Show splash text 
	WinSet, Transparent, 220, ahk_class AutoHotkeyGUI
	WinWaitClose, ahk_class AutoHotkeyGUI ; Wait until GUI has disappeared
	splashInputGuiEscape:  ; Pressing Esc, Enter, or Alt+F4 submits and returns 
	Gui, splashInput: Destroy
	Return
	splashInputGuiClose:
	splashInputButtonOK:
	Gui, splashInput: Submit
	Gui, splashInput: Destroy
	; Progress, off
	Return


/* 
This function is useful for minimalistic text input.
Call it by passing the number of rows, you wish to use, along with the text.
It will return the edit box text in the 'input' variable. 
*/