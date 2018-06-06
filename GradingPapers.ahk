#InstallKeybdHook
#SingleInstance, force

Menu, Tray, Icon, D:\Dropbox\Code\AHK\icons\G.ico
TrayTip, %A_ScriptName%, Script loaded
global toggle := 0
AutoReload()


/*


 /$$      /$$                           /$$
| $$  /$ | $$                          | $$
| $$ /$$$| $$  /$$$$$$   /$$$$$$   /$$$$$$$
| $$/$$ $$ $$ /$$__  $$ /$$__  $$ /$$__  $$
| $$$$_  $$$$| $$  \ $$| $$  \__/| $$  | $$
| $$$/ \  $$$| $$  | $$| $$      | $$  | $$
| $$/   \  $$|  $$$$$$/| $$      |  $$$$$$$
|__/     \__/ \______/ |__/       \_______/
                                           
                                           
                                           
Word
*/

#IfWinActive, ahk_class OpusApp
F3::
	oWord := ComObjActive("Word.Application") 
    oWord.Selection.StartOf(Unit:=4)  ; Select current paragraph in Word
	oWord.Selection.MoveEnd(Unit:=4)
	oWord.Selection.Cut  
	dir = %UserProfile%\Google Drev\Adjunkt\Grading\Templates
	file = %dir%\ß comment.rtf
	oWord.Selection.InsertFile(FileName:=file)
    oWord.Application.ActiveDocument.Bookmarks("paste_here").Select ; Select bookmark in Word
    send {Enter}
	oWord.Selection.Paste ; Paste cell value, overwriting bookmark	
	oWord.Application.ActiveDocument.Bookmarks("paste_here").Delete
	Return
:*:\new::
	oWord := ComObjActive("Word.Application").ActiveDocument.ActiveWindow
	oWord.NewWindow.View.Type := 6
	sleep 100
	WinMove, A,, 0,, 342,
	Send, ^!{Space}
	; oWord.View.Type := 6
	; Sleep 100
	; WinMove, A,, 342,, 1024,
	; oWord.View.Type := 3
	; Send, ^{Home}
	; Run, %A_ScriptDir%\wordColorFeedback.ahk
	Return 
:*:\rockwell18::
	wFont := ComObjActive("Word.Application").Selection.Font 
	wFont.Name := "Rockwell"
	wFont.Size := 18
	wFont.Color := 2895921
	; ComObjActive("Word.Application").Selection.MoveEnd(Unit:=2)  
	Return
F2::
	ClipSaved = %ClipboardAll%  ; Save clipboard
	Clipboard = ; Clear clipboard 
	send ^c ; Copy selected text 
	ClipWait, 2, 
	author := ComObjActive("Word.Application").Application.UserName  ; Get current author
	path = %A_ScriptDir%\Word\commentHeadings.txt ; Path to txt file 
	txtRadioOptions(path) ; Parse txt file, return %options% for splashUI radio 
	splashRadio("Choose comment heading", options)  ; Prompt for new heading 
	if choice <> ; Run only if user chose something 
	{
		if choice = New
			Run, %path%
		Else if choice = Other
		{
			splashText("Enter heading", 10, , Clipboard)   ; Prompt for heading (clipboard as default text)
			if input <> x				  ; If input not "x" (type x to cancel the comment) 
			{
				Clipboard = %ClipSaved%  ; Restore clipboard
				ComObjActive("Word.Application").Application.UserName := input ; input becomes new heading 
				splashText("Comment", 10)   ; Prompt for comment 
				if input <> x				  ; If input not "x" (type x to cancel the comment) 
					wComment(input)           ; Insert comment  
				ComObjActive("Word.Application").Application.UserName := author ; Restore original author 
			}
		}
		Else if choice contains ris,ros,J·kup,sp¯rgsmÂl,tip,uklart,husk,formatering,struktur
		{
			ComObjActive("Word.Application").Application.UserName := choice ; splashRadio choice above becomes new heading 
			splashText("Comment", 10)   ; Prompt for comment 
			if input <> x				  ; If input not "x" (type x to cancel the comment) 
				wComment(input)           ; Insert comment 
			ComObjActive("Word.Application").Application.UserName := author ; Restore original author 
		}
		Else {
			ComObjActive("Word.Application").Application.UserName := choice ; splashRadio choice above becomes new heading 
			wComment("")           ; Insert comment 
			ComObjActive("Word.Application").Application.UserName := author ; Restore original author 
		}
	}    
    Clipboard = %ClipSaved%  ; Restore clipboard
	Return

+#n:: ; Change user name (effectively changing the heading and color of a comment )
	author := ComObjActive("Word.Application").Application.UserName  ; Get current author
	splashText("Comment heading?", 2)  ; Prompt for new heading 
	if input <> ; Run only if user entered something 
		{
			ComObjActive("Word.Application").Application.UserName := input 
			splashText("Comment", 10)   ; Prompt for comment 
			if input <>
				wComment(input)           ; Insert comment 
			ComObjActive("Word.Application").Application.UserName := author ; Restore original author 
		}


^!l::wLink() ; Insert hyperlink

!+e::editComment() ; Edit last comment
	; oWord := ComObjActive("Word.Application")
	; oWord.ActiveDocument.Comments(oWord.ActiveDocument.Comments.Count).Edit
	; Return

F11:: ; Toggle distraction free mode 
	toggle++
	if  mod(toggle, 2) = 1
	{
		Send, !6
		Sleep 100
		Send, f
		Sleep 100
		Send, ^{Tab}
		Sleep 100s
		Send, !r
		Send, 34{Tab}
		Sleep 100
		Send, 26{Tab}
		Sleep 100
		Send, 15{Tab}
		Sleep 100
		Send, {Enter}
		Sleep 100
		oWord := ComObjActive("Word.Application")
		oWord.ActiveDocument.Range.Font.Color := 0x0098FE
		oWord.ActiveDocument.ActiveWindow.View.Zoom.Percentage := 160
		oWord.ActiveDocument.ActiveWindow.ToggleRibbon
	}
	if  mod(toggle, 2) = 0
	{
		oWord := ComObjActive("Word.Application")
		oWord.ActiveDocument.Range.Font.Color := -16777216
		oWord.ActiveDocument.ActiveWindow.View.Zoom.Percentage := 100
		oWord.ActiveDocument.ActiveWindow.ToggleRibbon
		Sleep 200
		Send, !6
		Sleep 200
		Send, i
	}
	Return

:*:\grade::
	InputBox, grade
	Send,Karakter: %grade% `rFejltyper:`r
	Send, !Y1
	Sleep, 200
	Send, {Down}{Right 2}{Enter}
	Sleep, 100
	Send, j
	Return

:*:\table::
	; oWord := ComObjActive("Word.Application") 
	myTable := oWord.Application.ActiveDocument.Tables
	myTable.Add(Range:=oWord.Selection.Range, NumRows:=5, NumColumns:=2).Columns(1).SetWidth(150,2).Font.Size := 14.Face := "Courier New"
	myTable.Item(1).Range.Font.Size := 14
	myTable.Item(1).Style := "Gittertabel 4 - farve 2"
	Add(1, 1, "Karakterer")
	Add(2, 1, "Skriftsprog:")
	splashText("Surface-level grade?", 1)
	Add(2, 2, input)
	splashText("Surface-level comment", 3)
	Add(3, 2, input)
	Add(4, 1, "Indhold:")
	splashText("Content-level grade?", 1)
	Add(4, 2, input)
	splashText("Content-level comment", 3)
	Add(5, 2, input)

	Add(row, col, txt)
	{
		global
		myTable.Item(1).Cell(row,col).Range.Text := txt
	}
	Return

/*
  /$$$$$$                                /$$                 /$$$$$$$$                                         
 /$$__  $$                              | $$                | $$_____/                                         
| $$  \__/  /$$$$$$   /$$$$$$   /$$$$$$ | $$  /$$$$$$       | $$     /$$$$$$   /$$$$$$  /$$$$$$/$$$$   /$$$$$$$
| $$ /$$$$ /$$__  $$ /$$__  $$ /$$__  $$| $$ /$$__  $$      | $$$$$ /$$__  $$ /$$__  $$| $$_  $$_  $$ /$$_____/
| $$|_  $$| $$  \ $$| $$  \ $$| $$  \ $$| $$| $$$$$$$$      | $$__/| $$  \ $$| $$  \__/| $$ \ $$ \ $$|  $$$$$$ 
| $$  \ $$| $$  | $$| $$  | $$| $$  | $$| $$| $$_____/      | $$   | $$  | $$| $$      | $$ | $$ | $$ \____  $$
|  $$$$$$/|  $$$$$$/|  $$$$$$/|  $$$$$$$| $$|  $$$$$$$      | $$   |  $$$$$$/| $$      | $$ | $$ | $$ /$$$$$$$/
 \______/  \______/  \______/  \____  $$|__/ \_______/      |__/    \______/ |__/      |__/ |__/ |__/|_______/ 
                               /$$  \ $$                                                                       
                              |  $$$$$$/                                                                       
                               \______/                                                                        
Google Forms
*/

#IfWinActive, Feedback: Surface Errors
; ~LButton::
;	If (A_TimeSincePriorHotkey<400) and (A_PriorHotkey="~LButton")
; +~LButton::
PgUp::
	{
  			Clipboard := 
  			Send ^a^c
  			If Clipboard is Integer
  			{
  				Clipboard := SubStr(Clipboard,1,10)+1
  				Send %Clipboard%
  				Return
		}
		Return
	}
Return
; ~RButton::
; 	If (A_TimeSincePriorHotkey<400) and (A_PriorHotkey="~RButton")
; ^~LButton::
PgDn::
	{
		Clipboard :=
		Send ^a^c
  			If Clipboard is Integer
  			{
  				Clipboard := SubStr(Clipboard,1,10)-1
  				Send %Clipboard%
  				Sleep 100
  				Send {Esc}
  				Return
		}
		Return
	}
Return

#IfWinActive, ahk_class OpusApp

/*
  /$$$$$$        /$$                               /$$       /$$                    
 /$$__  $$      | $$                              | $$      |__/                    
| $$  \ $$  /$$$$$$$ /$$    /$$ /$$$$$$   /$$$$$$ | $$$$$$$  /$$  /$$$$$$   /$$$$$$ 
| $$$$$$$$ /$$__  $$|  $$  /$$//$$__  $$ /$$__  $$| $$__  $$| $$ /$$__  $$ /$$__  $$
| $$__  $$| $$  | $$ \  $$/$$/| $$$$$$$$| $$  \__/| $$  \ $$| $$| $$$$$$$$| $$  \__/
| $$  | $$| $$  | $$  \  $$$/ | $$_____/| $$      | $$  | $$| $$| $$_____/| $$      
| $$  | $$|  $$$$$$$   \  $/  |  $$$$$$$| $$      | $$$$$$$/| $$|  $$$$$$$| $$      
|__/  |__/ \_______/    \_/    \_______/|__/      |_______/ |__/ \_______/|__/      
                                                                                    
                                                                                    
                                                                                    
Adverbier
*/
+!a::
	splashRadio("Adverbiers placering", "SmÂadverbier|Korte stedsadverbier|Lange stedsadverbier|Tidsadverbier")
	if choice <> 
		if choice = SmÂadverbier
		{
			wComment("SmÂadverbier placeres mellem subjekt og verballed eller efter f¯rste hjÊlpeverbum.")
			editComment()
			wNewLine()
			wLink("LÊs mere pÂ ENGRAM", "http://www.engram.dk/grammatik/ordstilling/#O41")
		}
		if choice = Korte stedsadverbier
		{			
			wComment("Korte stedsadverbialer placeres ofte efter hovedverbet (+ evt. objektet)")
			editComment()
			wNewLine()
			wLink("LÊs mere pÂ ENGRAM", "http://www.engram.dk/grammatik/ordstilling/#O44")
		}
		if choice = Lange stedsadverbier
		{	
			wComment("Lange stedsadverbialer placeres ofte enten f¯rst eller sidst i sÊtningen")
			editComment()
			wNewLine()
			wLink("LÊs mere pÂ ENGRAM", "http://www.engram.dk/grammatik/ordstilling/#O44")
		}
		if choice = Tidsadverbier
		{
			wComment("Tidsadverbier placeres oftest f¯rst eller sidst i sÊtningen.")
			editComment()
			wNewLine()
			wLink("LÊs mere pÂ ENGRAM", "http://www.engram.dk/grammatik/ordstilling/#O45")
		}
	Return
	
/*
    
 /$$   /$$                                                
| $$  /$$/                                                
| $$ /$$/   /$$$$$$  /$$$$$$/$$$$  /$$$$$$/$$$$   /$$$$$$ 
| $$$$$/   /$$__  $$| $$_  $$_  $$| $$_  $$_  $$ |____  $$
| $$  $$  | $$  \ $$| $$ \ $$ \ $$| $$ \ $$ \ $$  /$$$$$$$
| $$\  $$ | $$  | $$| $$ | $$ | $$| $$ | $$ | $$ /$$__  $$
| $$ \  $$|  $$$$$$/| $$ | $$ | $$| $$ | $$ | $$|  $$$$$$$
|__/  \__/ \______/ |__/ |__/ |__/|__/ |__/ |__/ \_______/
/                                                        
                                                          
                                                          

Komma

*/
+!,::
	wComment("")
	editComment()
	wLink("Der skal komma her. Hvorfor?", "http://www.engram.dk/grammatik/tegnsaetning/#T22")
	Send, {Esc}
	Return
:*:,that::
	wLink("Aldrig komma foran 'that'", "http://www.engram.dk/grammatik/tegnsaetning/#T21")
	Return

/*
                                                                                                                                                                                                                                                             
  /$$$$$$                                 /$$$$$$  /$$                                         /$$                    
 /$$__  $$                               /$$__  $$| $$                                        |__/                    
| $$  \__/  /$$$$$$  /$$$$$$$   /$$$$$$ | $$  \__/| $$  /$$$$$$  /$$    /$$ /$$$$$$   /$$$$$$  /$$ /$$$$$$$   /$$$$$$ 
| $$ /$$$$ /$$__  $$| $$__  $$ |____  $$| $$$$    | $$ /$$__  $$|  $$  /$$//$$__  $$ /$$__  $$| $$| $$__  $$ /$$__  $$
| $$|_  $$| $$$$$$$$| $$  \ $$  /$$$$$$$| $$_/    | $$| $$$$$$$$ \  $$/$$/| $$$$$$$$| $$  \__/| $$| $$  \ $$| $$  \ $$
| $$  \ $$| $$_____/| $$  | $$ /$$__  $$| $$      | $$| $$_____/  \  $$$/ | $$_____/| $$      | $$| $$  | $$| $$  | $$
|  $$$$$$/|  $$$$$$$| $$  | $$|  $$$$$$$| $$      | $$|  $$$$$$$   \  $/  |  $$$$$$$| $$      | $$| $$  | $$|  $$$$$$$
 \______/  \_______/|__/  |__/ \_______/|__/      |__/ \_______/    \_/    \_______/|__/      |__/|__/  |__/ \____  $$
                                                                                                             /$$  \ $$
                                                                                                            |  $$$$$$/
                                                                                                             \______/ 
Genaflevering
*/
+!d::
	splashText("Danish'ism - from which expression?", 2)
	if input <>
	{
		oWord := ComObjActive("Word.Application") 
		author := oWord.Application.UserName  ; Get current author
		oWord.Application.UserName := "Danishism"
		wComment(input)
		oWord.Application.UserName := author ; Restore original author 
	}
	Return
:*:\ingverb::
	wLink("Nogle verber f¯lges altid af ing-form", "http://www.engram.dk/grammatik/verber/#V92")
	Send, {Esc}
	Return
:*:\do::
	wText("Forklar hvornÂr man bruger do-omskrivning vha. ")
	wLink("dette afsnit pÂ ENGRAM", "http://www.engram.dk/grammatik/verber/#V7")
	Send, {Esc}
	Return
:*:\udvidet::
	wText("Forklar hvornÂr man bruger udvidet tid vha. ")
	wLink("dette afsnit pÂ ENGRAM", "http://www.engram.dk/grammatik/verber/#V4")
	Send, {Esc}
	Return
+!?::
	author := oWord.Application.UserName  ; Get current author
	oWord.Application.UserName := "Forklar fejlen"
	splashText("Kommentar", 2, , "Hvad er der galt her? Forklar i din genaflevering")
	wComment(input)
	oWord.Application.UserName := author ; Restore original author 
	Return
+!o::
	author := oWord.Application.UserName  ; Get current author
	oWord.Application.UserName := "Omformuleres"
	splashText("Kommentar", 2, , "OmformulÈr")
	wComment(input)
	oWord.Application.UserName := author ; Restore original author 
	Return
+!m::
	author := oWord.Application.UserName  ; Get current author
	oWord.Application.UserName := "Omformuleres"
	wComment("Dette skal omformuleres, sÂ det bliver mere klart")
	oWord.Application.UserName := author ; Restore original author 
	Return
+!k::
	author := oWord.Application.UserName  ; Get current author
	oWord.Application.UserName := "Omformuleres"
	wComment("Lidt klodset formuleret - omformulÈr")
	oWord.Application.UserName := author ; Restore original author 
	Return
+!h::
	author := oWord.Application.UserName  ; Get current author
	oWord.Application.UserName := "Omformuleres"
	wComment("SÊtningen her ender lidt brat. Skriv lidt mere eller omformulÈr")
	oWord.Application.UserName := author ; Restore original author 
	Return

	

/*

 /$$   /$$                                                             /$$                        
| $$  /$$/                                                            | $$                        
| $$ /$$/   /$$$$$$  /$$$$$$/$$$$  /$$$$$$/$$$$   /$$$$$$  /$$$$$$$  /$$$$$$    /$$$$$$   /$$$$$$ 
| $$$$$/   /$$__  $$| $$_  $$_  $$| $$_  $$_  $$ /$$__  $$| $$__  $$|_  $$_/   |____  $$ /$$__  $$
| $$  $$  | $$  \ $$| $$ \ $$ \ $$| $$ \ $$ \ $$| $$$$$$$$| $$  \ $$  | $$      /$$$$$$$| $$  \__/
| $$\  $$ | $$  | $$| $$ | $$ | $$| $$ | $$ | $$| $$_____/| $$  | $$  | $$ /$$ /$$__  $$| $$      
| $$ \  $$|  $$$$$$/| $$ | $$ | $$| $$ | $$ | $$|  $$$$$$$| $$  | $$  |  $$$$/|  $$$$$$$| $$      
|__/  \__/ \______/ |__/ |__/ |__/|__/ |__/ |__/ \_______/|__/  |__/   \___/   \_______/|__/      
                                                                                                                                                                                                    

Kommentar
*/
^Tab:: 
	oWord := ComObjActive("Word.Application") 
	wText("INDRYK>")
	Send, +{Home}
	oWord.Application.Selection.Range.HighlightColorIndex := 16 ; Grey (generelle kommentar)
	Return
:*:\danskeder::
	wLink("Se her, hvordan man oversÊtter det danske 'der' pronomen", "https://swipe.to/4430c")
	Return
+!-::wComment("Bindestreg imellem her")
+!u::wComment("Undlad artikler ved utÊllelige substantiver der angiver noget generelt, begrebsligt.")
+!g::
	Send, ^{Left}+^{Right}^c
	splashText("Word to define?", 2, , Clipboard)
	if input <> ; If user input not empty 
	{
		oWord := ComObjActive("Word.Application")
    	URL := % "https://www.google.dk/search?site=webhp&source=hp&q=define%20" . input
    	Text := "Check it out"
    	author := oWord.Application.UserName  ; Get current author
		oWord.Application.UserName := "Forkert ordvalg"
    	oWord.ActiveDocument.Comments.Add(oWord.Selection.Range)
	    oWord.ActiveDocument.Hyperlinks.Add(oWord.Selection.Range, URL,"","", text) 
		oWord.Application.UserName := author ; Restore original author 
	}
	Send, {Esc}
	Return
:*:\indryk::Husk at indrykke f¯rste linje i et afsnit, eller brug Words ‚Äùafstand efter afsnit‚Äù funktion, sÂ der er en klar skeln imellem afsnittene. 
:*:\ordi"::Det er ikke n¯dvendigt at bruge gÂse¯jne til ord, hvor du ikke vil udtrykke en anden mening and den Âbenbare.
:*:\tvtitler::Titler pÂ tv serier skrives med kursiv, imens navnene pÂ specifikke episoder skrives med "gÂse¯jne".
:*:\--::Her skulle du hellere have brugt en tankestreg. Kan du se forskellen`?{Esc}{Home}
; :*:\ec::{{}Esc{}}{{}Home{}}
+!r::wComment("Un¯dig gentagelse")
:*:\aleksandras::Du behersker et rigtig godt engelsk skriftsprog. Men du har skrevet et resume af en bestemt del af Vampire Diaries, hvor opgaven l¯d pÂ at beskrive den verden historien foregÂr i. 
:*:\ordklasse1::
	oWord := ComObjActive("Word.Application")
	author := oWord.Application.UserName  ; Get current author
	oWord.Application.UserName := "Forkert ordklasse"
	oWord.ActiveDocument.Comments.Add(oWord.Selection.Range)
	splashText("Forkert ordklasse?", 2)
	if input <> 
	{	
		o1 := input
		splashText("Korrekt ordklasse?", 2)
		if input <> 
		{	
		o2 := input
			wText("Du har skrevet ordet som et " . o1 . ", men skulle have skrevet det som et " . o2)
			wNewLine()
			wLink("Link til Ordklasser pÂ ENGRAM", "http://www.engram.dk/grammatik/ordklasser/")
			Send {Esc}{Home}
		}
	oWord.Application.UserName := author ; Restore original author 
	}
	Return
:*:\ordklasse2::
	splashText("Forkert ordklasse?", 2)
	o1 := input 
	splashText("Korrekt ordklasse?", 2)
	o2 := input
	wText("Du har skrevet disse som " . o1 . ", men skulle have skrevet dem som " . o2)
	wNewLine()
	wLink("Link til Ordklasser pÂ ENGRAM", "http://www.engram.dk/grammatik/ordklasser/")
	Send {Esc}{Home}
	Return

:*:\*prop::Ikke en del af propriet og skrives derfor ikke med smÂt
:*:\antons::men du bevÊger dig lidt ind pÂ at give et resume til sidst
:*:\sprog12::Godt sprog med fÂ grammatiske fejl
:*:\sprog10::Godt sprog med nogle grammatiske fejl
:*:\stavtal::Man staver oftest tal under 10 ("two" i stedet for "2", osv.)
:*:\sprog7::Dit sprog hÊmmes af dine ordvalg og grammatiske fejl
:*:\use::You "use" a tool but "spend" time.
:*:\economic::
	Send, _economic_: related to trade, industry or money.`r_economical_: not using a lot of money`r
	wLink("Read more here", "http://dictionary.cambridge.org/grammar/british-grammar/economic-or-economical")
	Return
+!s::
	Send, {F2}
	wLink("PÂ engelsk benyttes nÊsten kun ligefrem ordstilling", "http://www.engram.dk/grammatik/ordstilling/#ligefremogomvendt")
	Send, {Esc}{Home}
	Return
:*:\fejlplaceretadv::wLink("Fejlplaceret adverbium", "http://www.engram.dk/grammatik/ordstilling/#O41")
+!t::wComment("Stavefejls-forkert-ordvalg: slÂ bÂde dette og det danske ord, du havde i tankerne, op i en ordbog")
:*:\filmtitler::Filmtitler skrives med kursiv og ikke gÂse¯jne	
+!f::
	oWord := ComObjActive("Word.Application") 
	oWord.ActiveDocument.TrackRevisions := False
	oWord.Application.Selection.Font.StrikeThrough := True
	author := oWord.Application.UserName  ; Get current author
	oWord.Application.UserName := "Overfl¯digt"
	wComment("")
	oWord.Application.UserName := author ; Restore original author
	Return
:*:(relativs::parentetisk relativsÊtning{Esc}{Home}
+!l:: ; Prompt & indsÊt "brug heller" kommentar
	splashText("Brug heller hvad?", 3)
	if input <> 
	{
		wComment("Brug heller: ")
		editComment()
		Send, ^i
		wText(input)
		Send, {Esc}
	}
	Return	
+!b::wComment("Ikke helt forkert, men kan du lave et bedre ordvalg?")
:*:\jobtitler::Jobtitler skrives med Stort Begyndelsesbogstav
:*:\meremund::Du skal vÊre bedre til at udnytte timerne til at ¯ve dit mundtlige engelsk
:*:grammatiskkorrekt::sÂ den bliver grammatisk korrekt 
:*:\lastsentence::Pr¯v at afslut din essay med en velformuleret sÊtning, der giver din lÊser stof til eftertanke
:*:\:citat::Formelt introducerede citater skal have et forudgÂende kolon. NÂr citatet er skrevet ind i teksten, skal du enten undlade forudgÂende tegnsÊtning eller bruge komma
:*:\kompass::NÂr kompasretningerne refererer til Èt geografisk omrÂde eller region skrives de oftest med stort. 
:*:.And::VÊr sparsommelig med at starte en sÊtning med en konjunktion. Kan du omformulere denne og den foregÂende sÊt-ning?
:*:\Stort::Hvorfor Med Stort Begyndelsesbogstav?
:*:\hanging::Dette hÊnger lidt for sig selv. Hvad mangler der?
:*:\hvad::Jeg forstÂr ikke hvad du mener. Pr¯v at omformulere
:*:\andenprp::Hvilken anden prÊposition skal du bruge her?
:*:\spring::Undlad at skifte imellem verbernes tider i samme afsnit
:*:engram,::http://www.engram.dk/grammatik/tegnsaetning/{#}T22 
:*:\titlerikursiv::Titler pÂ ting, der betragtes som en helhed (en bog, navnet pÂ et -blad osv.) skrives i kursiv
