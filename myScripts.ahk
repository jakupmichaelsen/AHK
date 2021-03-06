#NoEnv ; Significantly improves performance by not looking up empty variables as environmental variables 
#NoTrayIcon
#SingleInstance, force
AutoReload()

options = 
    ( LTrim Join|
        Adjunkt
        adobeAcrobat.ahk
        Annotations.ahk
        AutoHotkey.ahk
        chromePresentationMode.ahk
        colorFeedback.ahk
        docx2pdf.ahk
        docx2xl.ahk
        feedbackTemplate2docx.ahk
        filenames2Excel.ahk
        functionKeys.ahk
        googleComments.ahk
        googleNgramViewer.ahk
        IE Element Spy.ahk
        lectioFilenames2Excel.ahk
        lectioLinks.ahk
        mailMerge.ahk
        myScripts.txt
        pdf2docx.ahk
        removeHeadersFooters.ahk
        removeAllTablesInWord.ahk
        removeLectioIDs.ahk
        renameSubDirPDFs.ahk
        reveal-md.ahk
        Rubric.ahk
        runningScripts.ahk
        splashUI_image.ahk
        test.ahk
        TextExpansions.ahk
        wBookmarks.ahk
        WinTrigger.ahk
        wordFeedback.ahk
        workFlowyPresentationMode.ahk
        xlFilenameRenamer.ahk
    )

splashList("", options, 1, 13)
    if choice <> ; Run only if choice isn't empty 
        Run, %A_ScriptDir%\%choice%
    Return