#SingleInstance, force 
IniRead, ahk, splashUI.ini, paths, ahk

dir = %ahk%\images

splashDir(dir)
splashList("Splash images", items)
if choice <>
	SplashImage = %dir%\%choice%

#½::
	Toggle++
	if  mod(Toggle, 2) = 1 
		splashImageGUI(SplashImage, "Center", "Center", 0xa349a4, true)
	Else 
		Gui, XPT99:Destroy
	Return
