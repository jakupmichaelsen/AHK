readyHighlighterTool(){ ; Function for Adobe Acrobat DC highliger window
    Send ^e
    WinWait, Highlighter Tool Properties ahk_class AVL_AVFloating
    WinActivate, Highlighter Tool Properties ahk_class AVL_AVFloating
    Sleep 200
    Click
    Send {Down 7}
}
