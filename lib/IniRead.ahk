IniRead(){
	global
	IniRead, background, %A_ScriptDir%\splashUI.ini, colors, background
	StringTrimLeft, background, background, 1
	IniRead, header_color, %A_ScriptDir%\splashUI.ini, colors, header_color
	StringTrimLeft, header_color, header_color, 1
	IniRead, font_color, %A_ScriptDir%\splashUI.ini, colors, font_color
	StringTrimLeft, font_color, font_color, 1
	IniRead, header, %A_ScriptDir%\splashUI.ini, fonts, header
	IniRead, font, %A_ScriptDir%\splashUI.ini, fonts, font
	IniRead, typingFont, %A_ScriptDir%\splashUI.ini, fonts, typingFont
	IniRead, size, %A_ScriptDir%\splashUI.ini, fonts, size
	IniRead, style, %A_ScriptDir%\splashUI.ini, fonts, style
}