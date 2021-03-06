#SingleInstance, force 
#Persistent
AutoReload()
#NoTrayIcon

IniRead, onedrive, splashUI.ini, paths, onedrive
global mdScreenshots
mdScreenshots = %onedrive%\Markdown\screenshots 


WatchFolder(mdScreenshots, "ReportFunction")
ReportFunction(Directory, Changes) { 
	For Each, Change In Changes {
		Action := Change.Action
		Path := Change.Name
		Name := SubStr(StrReplace(Path, Directory), 2) ; Retrieve all, save the first, string of Path minus Directory 

		; Action 1 (added) = File gets added in the watched folder
		If (Action = 1)
		{
			splashNotify("Screenshot path copied to clipboard") 
			Sleep 1000
			Clipboard = `!`[%Name%`]`(screenshots\%Name%`)
		}
	}
}

