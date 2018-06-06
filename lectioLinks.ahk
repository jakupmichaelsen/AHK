#InstallKeybdHook
#SingleInstance, force
#Hotstring, * ?
#NoTrayIcon
SendMode Input
AutoReload()
browser = C:\Program Files (x86)\Google\Chrome\Application\chrome.exe  --profile-directory=Default

LectioTop:
splashList("Lectio links", "Mit skema|Forside|2h|2r|1f|AP|Avanceret skema|Annex skema|Lokaler lige nu|Lærerskema", 0)

if choice <>
{
  if choice = Forside
  {
  	run %browser% --app=https://www.lectio.dk/lectio/248/forside.aspx
  	Sleep, 2500
  	Send {Enter}
  }
  Else if choice = Lærerskema
  	run %browser% --app=https://www.lectio.dk/lectio/248/FindSkema.aspx?type=laerer
  Else if choice = Lokaler lige nu
  	run %browser% --app=https://www.lectio.dk/lectio/248/SkemaAvanceret.aspx?type=aktuelleallelokaler
  Else if choice = Mit skema
    run %browser%  --app=https://www.lectio.dk/lectio/248/SkemaNy.aspx?type=laerer&laererid=9130528370,,, LectioPID
  Else if choice = Avanceret skema 
    run %browser% --app=https://www.lectio.dk/lectio/248/FindSkemaAdv.aspx
  Else if choice = Annex skema
  {
    URL = "https://www.lectio.dk/lectio/248/SkemaAvanceret.aspx?type=skema&studentsel=&teachersel=&klassesel=&holdsel=&holdmedsel=&lokalesel=&holdfaggrpsel=&teacherfaggrpsel=&ressel=9203917380,9203918459,9203919422,9203920428"
    run %browser% --app=%URL%
  }
  
  Else {
  	classChoice := choice 
  	splashList("Lectio links", "Aktiviteter|Elever|Klasseskema|Studieplan|Opgaver", 0)
  	if choice <>
  	{
	    if choice = Elever 
	      klasseID(classChoice)
	    Else if choice = Klasseskema
	      klasseID(classChoice)
	    Else if choice = Studieplan 
	      holdID(classChoice)
	    Else 
	      holdID(classChoice)
	      lectioLinks(choice, klasse)
	      run, %browser% --app=%link%
  	}
  	Else
    	Gosub, LectioTop
	}
	ExitApp
}

lectioLinks(choice, klasse){
global
if choice = Aktiviteter 
  link = https://www.lectio.dk/lectio/248/aktivitet/AktivitetListe2.aspx?holdelementid=%klasse%
else if choice = Elever
  link = https://www.lectio.dk/lectio/248/subnav/members.aspx?showstudents=1&reporttype=classpicture&klasseid=%klasse%
else if choice = Klasseskema
  link = https://www.lectio.dk/lectio/248/SkemaNy.aspx?type=stamklasse&klasseid=%klasse%
else if choice = Studieplan
  link = https://www.lectio.dk/lectio/248/studieplan.aspx?displaytype=ugeteksttabel&holdelementid=%klasse%
else if choice = Opgaver
  link = https://www.lectio.dk/lectio/248/opgave.aspx?holdelementid=%klasse%
Return
}
holdID(choice){
  global 
  if classChoice = 2h
  	klasse = 15630548864
  else if classChoice = 2r
  	klasse = 15611368392 
  else if classChoice = 1f
  	klasse = 20657714891
  else if classChoice = AP
  	klasse = 20369388345
  Return 
}
klasseID(classChoice){
  global 
  if classChoice = 2h
    klasse = 15589354317
  else if classChoice = 2r
    klasse = 15589358388 
  else if classChoice = 1f
	 klasse = 20403163920
  else if classChoice = AP
	 klasse = 20180217951
Return 
}
