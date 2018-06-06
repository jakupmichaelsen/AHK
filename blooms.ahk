#SingleInstance, force
niveauer = &Skabe|&Vurdere|&Analysere|&Anvende|&Forstå|&Huske

Skabe = &designe|&konstruere|&planlægge|&producere|&opfinde|&udforme|&innovere|&formulere|&udvikle
Vurdere = &efterprøve|&hypotisere|&kritisere|&eksperimentere|&bedømme|&teste|&argumentere
Analysere = &sammenligne|&organisere|&dekonstruere|&eksperimentere|&finde|&påvise|&strukturere|&stille spøsrgsmål ved
Anvende = &udvælge|&demonstrere|&illustrere|&udforme|&bruge|&løse opgaver|&dramatisere
Forstå = &tolke|&sammenfatte|&udlede|&omskrive|&beskrive|&klassificere|&forklare|&eksemplificere|&relatere
Huske = &genkende|&opliste|&beskrive|&identificere|&genkalde|&navngive|&lokalisere|&finde

splashUI("r", "Vælg taksonomisk niveau", niveauer)
if choice <>
{
	niveau := choice 
	
	StringReplace, niveauer, niveauer, &,,All
	Loop, Parse, niveauer, |, 
	{
		if niveau = %A_LoopField%
			splashUI("r", niveau, %A_LoopField%)
	}
	if choice <>
	{
		send %choice%
	}
}