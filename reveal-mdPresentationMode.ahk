#SingleInstance, force
#Hotstring, * ?
#InstallMouseHook
#SingleInstance, force 
AutoReload()

; #IfWinActive, reveal.js - Slide Notes – Google Chrome
PgDn::
	WinActivate, reveal.js - Slide Notes – Google Chrome
	WinWaitActive, reveal.js - Slide Notes – Google Chrome
	Send {Left}
	Return
PgUp::
	WinActivate, reveal.js - Slide Notes – Google Chrome
	WinWaitActive, reveal.js - Slide Notes – Google Chrome
	Send {Right}
	Return 



; Browser_Back::Send {Up}
; Browser_Forward::Send {Down}


