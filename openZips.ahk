#SingleInstance, force 
#Persistent
#NoTrayIcon
IniRead, user, splashUI.ini, paths, user 
AutoReload()
folder = %user%\Downloads

WatchFolder(folder, "ReportFunction")

ReportFunction(Directory, Changes) { 
	SetWorkingDir, %Directory%
	For Each, Change In Changes {
		Action := Change.Action
		Path := Change.Name

		; Action 1 (added) = File gets added in the watched folder
	   If (Action = 4) ; Renamed filed (downloading files have a temporary filename)
	   {
	   		IfInString, Path, .zip
	   		{
				filename := StrReplace(Path, Directory) ; Get the file name 
				StringTrimRight, zippath, Path, 4 ; Path to unzipped folder 
				StringTrimLeft, filename, filename, 1 ; Remove slash from filename 
				filename = `"%filename%`" ; Add quotation marks (for filenames with spaces)
	   		 	RunWait, %comspec% /c "7z x -o* %filename%" ; unzip
	   		 	Run, %zippath%
	   		 	FileDelete %path%
	   		}
	   }
	}
}
