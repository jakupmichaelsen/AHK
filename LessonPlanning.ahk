#InstallKeybdHook
#SingleInstance, force
SendMode Input
IniRead, ahk, splashUI.ini, paths, ahk

TrayTip, %A_ScriptName%, Script loaded

txtRadioOptions("%ahk%\lessonPlanning\AFM.md")
splashUI("r", "Faglige mål", options)
SendMode, Play 
if choice <>
	Send %choice%




:*:\afm::
	lessonPlanningTags("A", "FM") ; A-niveau's faglige mål
	Return
:*:\aks::
	lessonPlanningTags("A", "KS") ; A-niveau's kernestof
	Return
:*:\bfm::
	lessonPlanningTags("B", "FM") ; B-niveau's faglige mål
	Return
:*:\bks::
	lessonPlanningTags("B", "KS") ; B-niveau's kernestof
	Return


:*:\blooms::
	run, bloomsVerbs.ahk
	Return

/*
:*:\taxonomy::
	text := "Knowledge: Exhibit memory of learned materials by recalling facts, terms, basic concepts and answers`r`rComprehension: Demonstrate understanding of facts and ideas by organizing, comparing, translating, interpreting, giving descriptions, and stating the main ideas`r`rApplication: Using acquired knowledge. Solve problems in new situations by applying acquired knowledge, facts, techniques and rules in a different way`r`rAnalysis: Examine and break information into parts by identifying motives or causes. Make inferences and find evidence to support generalizations`r`rEvaluation: Present and defend opinions by making judgments about information, validity of ideas or quality of work based on a set of criteria`r`rSynthesis: Builds a structure or pattern from diverse elements; it also refers the act of putting parts together to form a whole (Omari, 2006). Compile information together in a different way by combining elements in a new pattern or proposing alternative solutions"
	Progress, show B1 x0 w400 h1000 zy20 c00 fs12 zh0, %text%,,, Courier New 
	splashUI("r", "Taxonomic level of cognitive domain", "Knowledge|Comprehension|Application|Analysis|Evaluation|Synthesis")
	Progress, off
	Send, %choice%
	Return

