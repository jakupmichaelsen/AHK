grades = 12: fremragende|10: fortrinligt|7: godt|4: jævnt|02: tilstrækkeligt|00: utilstrækkeligt|-3: ringe 

top:
splashList("gradeStrings", grades)

if choice <> 
{
	if choice = 12: fremragende
		send fremragende
	if choice = 10: fortrinligt
		send fortrinligt
	if choice = 7: godt
		send godt
	if choice = 4: jævnt
		send jævnt
	if choice = 02: tilstrækkeligt
		send tilstrækkeligt
	if choice = 00: utilstrækkeligt
		send utilstrækkeligt
	if choice = -3: ringe 
		send ringe 
}
Return 
