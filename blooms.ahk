#SingleInstance, force
niveauer = &Skabe|&Vurdere|&Analysere|&Anvende|&Forst�|&Huske

Skabe = &designe|&konstruere|&planl�gge|&producere|&opfinde|&udforme|&innovere|&formulere|&udvikle
Vurdere = &efterpr�ve|&hypotisere|&kritisere|&eksperimentere|&bed�mme|&teste|&argumentere
Analysere = &sammenligne|&organisere|&dekonstruere|&eksperimentere|&finde|&p�vise|&strukturere|&stille sp�srgsm�l ved
Anvende = &udv�lge|&demonstrere|&illustrere|&udforme|&bruge|&l�se opgaver|&dramatisere
Forst� = &tolke|&sammenfatte|&udlede|&omskrive|&beskrive|&klassificere|&forklare|&eksemplificere|&relatere
Huske = &genkende|&opliste|&beskrive|&identificere|&genkalde|&navngive|&lokalisere|&finde

splashUI("r", "V�lg taksonomisk niveau", niveauer)
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