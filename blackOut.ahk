#SingleInstance, force 
AutoReload()
TrayTip, %A_ScriptName%, Script loaded
; Turn Monitor Off 
; (2 = off, 1 = standby, -1 = on)

SendMessage, 0x112, 0xF170, 2,, Program Manager   