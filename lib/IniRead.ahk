IniRead(){
	global
	iniDir = D:\GitHub\AHK

; === paths === 
	IniRead, browser, %iniDir%\splashUI.ini, paths, browser
	IniRead, onedrive, %iniDir%\splashUI.ini, paths, onedrive
	IniRead, ahk, %iniDir%\splashUI.ini, paths, ahk
	IniRead, user, %iniDir%\splashUI.ini, paths, user
	IniRead, git, %iniDir%\splashUI.ini, paths, git

; === [colors ===
	IniRead, background, %iniDir%\splashUI.ini, colors, background
	IniRead, font_color, %iniDir%\splashUI.ini, colors, font_color
	IniRead, header_color, %iniDir%\splashUI.ini, colors, header_color
	; Trim "#"":
	StringTrimLeft, background, background, 1
	StringTrimLeft, header_color, header_color, 1
	StringTrimLeft, font_color, font_color, 1

; === fonts ===
	IniRead, header, %iniDir%\splashUI.ini, fonts, header
	IniRead, font, %iniDir%\splashUI.ini, fonts, font
	IniRead, typingFont, %iniDir%\splashUI.ini, fonts, typingFont
	IniRead, size, %iniDir%\splashUI.ini, fonts, size
	IniRead, big, %iniDir%\splashUI.ini, fonts, big
	IniRead, style, %iniDir%\splashUI.ini, fonts, style
}