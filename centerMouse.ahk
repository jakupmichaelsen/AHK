#SingleInstance, force
#Hotstring, * ?
#InstallMouseHook
#NoTrayIcon
AutoReload()
SysGet, mon2, Monitor, 2 ; Get info on external monitor 
mon2centerX := mon2left + ((mon2Right - mon2Left) /2 ) ; Calculate center 
mon2centerY := mon2Bottom / 2
#LButton::MouseMove, (A_ScreenWidth // 2), (A_ScreenHeight // 2) ; Centers mouse on primary display 
#RButton::MouseMove, mon2centerX, mon2centerY ; Centers mouse of expanded desktop display 
