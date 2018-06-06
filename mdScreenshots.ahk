#SingleInstance, force 
#Persistent
AutoReload()
#NoTrayIcon


WatchFolder("D:\ONEDRI~1\markdown\screenshots", "ReportFunction")

ReportFunction(Directory, Changes) { 
	For Each, Change In Changes {
		Action := Change.Action
		Path := Change.Name
		Name := SubStr(StrReplace(Path, Directory), 2) ; Retrieve all, save the first, string of Path minus Directory 

		; Action 1 (added) = File gets added in the watched folder
		If (Action = 1)
		{
			Clipboard = `!`[%Name%`]`(screenshots\%Name%`)
			TrayTip, Screenshot, Path copied to clipboad, 2, 1
		}
	}
}
