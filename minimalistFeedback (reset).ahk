#NoEnv ; Significantly improves performance by not looking up empty variables as environmental variables 
#SingleInstance, force

Try ComObjActive("Word.Application").ActiveDocument.Save
Sleep 500
Run, taskkill /F /IM winword*,,Hide
Sleep 500
Run, minimalistFeedback.ahk