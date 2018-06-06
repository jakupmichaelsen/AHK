#SingleInstance, force
#NoTrayIcon
AutoReload()

browser = C:\Program Files\Google\Chrome\Application\chrome.exe
url = fcp://@mail.silkeborg-gym.dk`,`%238850259/

F6::
path = %A_ScriptDir%\FirstClass links.txt ; Path to txt file 
txtRadioOptions(path) ; Function to parse txt file, readying if for splashUI radio 
splashUI("r", "directFirstClass", options)
if choice <>
{
	if choice = Pædagogikum
		link = %url%P`%C3`%A6dagogikum_J`%C3`%A1kup`%20Dahl`%20Michaelsen_l
	else if choice = Skolekom
		link = %browser% --app=https://web.skolekom.dk/
	else if choice = 3i EN
		link = %url%13i`%20EN
	else if choice = en_l
		link = %url%en_l
	else if choice = 3i_l
		link = %url%13i_l
	else if choice = Mail
		link = %url%MailBox 
	else if choice = Fra ledelse og kontor
		link = %url%Fra`%20ledelse`%20og`%20kontor_l
	else if choice = Fra kollega til kollega
		link = %url%Fra`%20kollega`%20til`%20kollega_l
}

Run %link%

Return
