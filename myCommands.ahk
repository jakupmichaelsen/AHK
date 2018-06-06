#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
#NoTrayIcon
IniRead, browser, splashUI.ini, paths, browser
IniRead, onedrive, splashUI.ini, paths, onedrive
IniRead, user, splashUI.ini, paths, user
SendMode Input
AutoReload()

path = %A_ScriptDir%\commandList.txt
txtList(path)

commandListTop:
splashList("", options)

if choice <> 
{
	if choice = Netprøver (folder)
		run D:\OneDrive - Silkeborg Gymnasium\Adjunkt\Eksamen\Censor\Skr. Censor\Netproever
	if choice = wBookmarks.ahk
		run %A_ScriptDir%\wBookmarks.ahk
	if choice = remHeadersFooters.ahk 
		run %A_ScriptDir%\remHeadersFooters.ahk
	if choice = removeAllTablesInWord.ahk
		run %A_ScriptDir%\removeAllTablesInWord.ahk
	if choice = lectioFilenames2Excel.ahk
		run %A_ScriptDir%\lectioFilenames2Excel.ahk
	if choice = removeLectioIDs.ahk
		run %A_ScriptDir%\removeLectioIDs.ahk
	if choice = workFlowyPresentationMode.ahk
		run %A_ScriptDir%\workFlowyPresentationMode.ahk
	if choice = xlFilenameRenamer.ahk
		run %A_ScriptDir%\xlFilenameRenamer.ahk
	if choice = Get Cohesive!
		run %onedrive%\Adjunkt\Metode\Writing\Get Cohesive!.pdf 
	if choice = minimalistFeedback.ahk
		run %A_ScriptDir%\minimalistFeedback.ahk
	if choice = splashUI_image.ahk
		run %A_ScriptDir%\splashUI_image.ahk
	if choice = splashUI_Image
		run splashUI_Image.ahk
	if choice = Todo 
		run D:\OneDrive - Silkeborg Gymnasium\Adjunkt\Todo.md
	if choice = Ngram
		run %onedrive%\Code\ahk\googleNgramViewer.ahk 
	if choice = Messenger
		run %browser% --app=https://www.messenger.com/t/kristina.thastrup
	if choice = Dr.dk
		run %browser% --app=https://www.dr.dk/tv
	if choice = Screen updating (Word)
		ComObjActive("Word.Application").Application.ScreenUpdating := True
	if choice = Work (Chrome)
		run %browser% --profile-directory="Profile 1"
    if choice = Jákup (Chrome)
		run %browser% --profile-directory="Default"
    if choice = Adjunkt
        run %onedrive%\Adjunkt
	if choice = Q10
        run %onedrive%\Apps\Q10\Q10.exe
    if choice = reveal-md 
		Gosub, start_reveal-md
	if choice = Login: Gyldendals Røde
		run https://loginconnector.gyldendal.dk/providers/ekey/login?accepturl=https`%3a`%2f`%2floginconnector.gyldendal.dk`%3a443`%2f`%2fNavigator`%2fSubmitLoginSuccess&declineurl=http`%3a`%2f`%2fordbog.gyldendal.dk`%2f`%2fapi`%2fsignin-gyldendal`%3fstate`%3dOkBMhGVjbz1TYiZ0sk0yQ2I33PIWgzoG6lxYEo78fckYyj1NefFDd-JvnUX7SavOe1erM6LlLvUojjweGcQBKmiwk3H-PeLVp5FfCBdduotL1jT6s6slw7mMRdXZw0yBVYCn2f5_zAfsa443l90CiE4DpIHVHA_bP22kyakmokht3l2Dq2UyFeROeNBMFdb5w7J2Xks5KRNjTqQIGLPA4eB6BKqRxClkciZgPH0pZ6lJCG_8Rxfo5P38FILZAESynIji__16eWI7k6VMXb3FibfZ__6n9qcTK6MWWotIOjd2CXZNSfWJrDVPLyPtIUyOUI2SgugAspUR9NkGBX59EYssm-LBBJZQdQGBLJN6SJCdpNBeKL1WM0IpOI2yXfW6_cEXuwPX15WZL95Znq-8tA&clientWebSite=Ordbog&clientCallingUrl=http`%3a`%2f`%2fordbog.gyldendal.dk`%2f#/
	if choice = Grupper
		run %browser% --app=https://wordwall.co.uk/user/209815/jdm/ --profile-directory="Profile 3"
	if choice = Google Play Books
		run %browser% --app=https://play.google.com/books
	if choice = Music 
		run %browser% --app=https://open.spotify.com/browse/featured
	if choice = Tampermonkey
		run %browser% --app=chrome-extension://dhdgffkkebhmkfjojejmpbldmpobfkfo/options.html#url=&nav=dashboard
	if choice = TTS
		ComObjCreate("SAPI.SpVoice").Speak(Clipboard)
	if choice = Note 
	{
		splashNoteSmall()
		Clipboard := input 
	}
	if choice = Kantinebrikpåfyldning
		run %browser% --app=https://services.pos-consult.dk/Public/1113/Auth/Login
	if choice = Stylish
		run %browser% --app=chrome-extension://fjnbnpbmkenffdnngjfgmeleoegfcffe/manage.html
	if choice = Recap
		run %browser% --app=https://app.letsrecap.com/teacher/#/class/88146/home
	if choice = Calendar
		run %browser% --app=https://calendar.google.com/calendar/render
	if choice = URL.html
	{
		StringReplace, Clipboard, Clipboard, http://
		StringReplace, Clipboard, Clipboard, https://
		StringReplace, Clipboard, Clipboard, /, /%A_Space%
		splashText("url.html", 1,"'x' = run without changes", Clipboard)
		if input <>
		{
			if input <> x ; 
			{
				html = <style> body { word-wrap: break-word; font-size: 300px; font-weight: 900; font-family: Rockwell;}</style>`r
				FileDelete, url.html
				FileAppend, %html%, url.html
				FileAppend, %input%, url.html
				run url.html
			}
			Else
				run url.html
		}
	}
	if choice = large.html
	{
		splashText("large.html", 10,"'x' = run without changes", Clipboard)
		if input <>
		{
			if input <> x ; 
			{
				html = <style> body { font-size: 200px; font-family: Rockwell;}</style>`r
				FileDelete, large.html
				FileAppend, %html%, large.html, UTF-16
				FileAppend, %input%, large.html, UTF-16
				run large.html
			}
			Else
				run large.html
		}
	}
	if choice = Screen off 
		SendMessage, 0x112, 0xF170, 2,, Program Manager   
	if choice = Grades
		Gosub, gradeStrings
	if choice = Kolleger
		run %browser% --profile-directory=Default --app=http://www.gymnasiet.dk/personalet/
	if choice = Kørselsordningen
		run http://www.glavind-groenne.dk/koersel/index.php
	if choice = Kristina
		run %browser% --profile-directory="Profile 2"
	if choice = FirstClass 
		run %user%\AppData\Local\Temp\fcctemp
	if choice = Chrome
		run %browser% --profile-directory=Default
	if choice = SG Chrome
		run %browser% --profile-directory="Profile 1"
	if choice = Run Script
		Gosub, runScripts
	if choice = Close Script
		Gosub, closeScripts
	if choice = Coggle
		run %browser% --profile-directory=Default --app=https://coggle.it/
	if choice = Acrobat
		run %PROGRAMFILES%\Adobe\Acrobat DC\Acrobat\Acrobat.exe
	if choice = Word
		run %PROGRAMFILES%\Microsoft Office\Root\Office16\WINWORD.EXE
	if choice = Excel
		run %PROGRAMFILES%\Microsoft Office\root\Office16\EXCEL.EXE
	if choice = PowerPoint 
		run %PROGRAMFILES%\Microsoft Office\Root\Office16\POWERPNT.EXE
    if choice = POWER 
        gosub POWER 
    if choice = WIFI 
        gosub WIFI 
    if choice = CMD
    	Gosub, CMD
    if choice = Paint
        run %windir%\system32\mspaint.exe
    if choice = Explorer
        run %user%\AppData\Roaming\Microsoft\Windows\Recent
    if choice = Desktop
        run %onedrive%\Desktop
    if choice = Downloads
        run %user%\Downloads
    if choice = Google
        run %user%\Google Drev
    if choice = Apps 
        run %onedrive%\Apps
    if choice = Markdown
        run %onedrive%\Markdown
    if choice = OneDrive
    ; {
	; 	dir = %onedrive%\
    ;     splashDir(dir, ,1)
    ;     splashList("onedrive folders", items)
    ;     if choice <> ; Run only if choice isn't empty
    ;        Run, %dir%\%choice%
    ;     Else
    ;         Gosub, commandListTop
    ; }
    	run %onedrive%
    if choice = AHK
        run %onedrive%\Code\ahk
}
Return 

start_reveal-md: ; Start markdown presentation 
    path = %onedrive%\Markdown\
    splashDir(path, "*.md")
    splashList("Select markdown to present", items)
    if choice <> ; Run only if choice isn't empty 
    	Run, %comspec% /k "cd %path% && reveal-md %choice%"
	; {
	; 	Gosub, CMD
	; 	WinWaitActive, Windows PowerShell ahk_class ConsoleWindowClass
	; 	Send, cd %path%{Enter}
	; 	Sleep 50
	; 	Send, reveal-md %choice% {Enter} 
	; }
    Return


runScripts:  ; Run scripts
    path = %A_ScriptDir%\myScripts.txt ; Path to txt file 
    txtList(path) ; Function to parse txt file, readying if for splashUI 
    splashList("Select script to run", options)
    if choice <> ; Run only if choice isn't empty 
        Run, %A_ScriptDir%\%choice%
    Return

closeScripts:  ; Run scripts
    splashText("Close script", 1, "")
    if input <> ; if input is not empty
	{
		DetectHiddenWindows, On
		WinClose, %A_ScriptDir%\%input% ahk_class AutoHotkey
	}
    TrayTip, AutoHotkey, %input% closed
    Return

gradeStrings: ; Send grade words 
	grades = 12: fremragende|10: fortrinligt|7: godt|4: jævnt|02: tilstrækkeligt|00: utilstrækkeligt|-3: ringe 
	splashList("gradeStrings", grades, sorted)
	if choice <> 
	{
		if choice = 12: fremragende
			send fremragende
		if choice = 10: fortrinligt
			send meget god
		if choice = 7: godt
			send god
		if choice = 4: jævnt
			send middelmådig
		if choice = 02: tilstrækkeligt
			send tilstrækkelig
		if choice = 00: utilstrækkeligt
			send utilstrækkelig
		if choice = -3: ringe 
			send ringe 
	}
	Return 

POWER: ; Toggle power saving and performance power options
    toggle++
    if  mod(toggle, 2) = 1
        run %ComSpec% /c powercfg -setactive a1841308-3541-4fab-bc81-f71556f20b4a,, Hide
    if  mod(toggle, 2) = 0
        run %ComSpec% /c powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c,, Hide
    Return

WIFI: ; Repair wireless connection
    Send, {Enter}
    Sleep 300
    MouseClick, right, 1240, -10
    Sleep 100
    SendPlay {Click 1240, 10}
    Return

CMD:
    ; IfWinActive, ahk_class CabinetWClass ; If Windows explorer active
    ; {
    ;     explorerPath()
    ;     run LexikosCMD.ahk
    ;     WinWaitActive, AHK CMD
    ;     Send cd %path%{Enter}cls{Enter}
    ; }
    ; Else
    ;     run CMD

	Send {LWin}
	WinWaitActive, ahk_class Windows.UI.Core.CoreWindow
	Send windows powershell
	Sleep 500
	Send {AppsKey}
	Sleep 100
	Send {Down}{Enter}
	; run "%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe", %HOMEDRIVE%%HOMEPATH%
    Return
