;==================預留功能==================
^+T::
    Return

;==================判斷銷單狀態（執行操作中）==================
^+Y::
    ToolTip, 正在執行選單操作..., 900, 300
    Send, {F10}l11
    Sleep, 100
    Return

;==================判斷銷單狀態（重新開啟銷單中）==================
^+U::
    ToolTip, 重新開啟銷單畫面中....., 900, 300
    WinGetText,Str,A
    Haystack := Str
    Needle := "功能快捷視窗"
    Atr := InStr(Haystack, Needle)
    while not (Atr = 1) {
        Sleep % 100
        Send, {F12}
        Sleep % 100
        break
    }
    WinGetText,Str,A
    Haystack := Str
    Needle := "銷貨單"
    Atr1 := InStr(Haystack, Needle)
    while not (Atr1 = 1) {
        Sleep % 100
        Send, !
        Send, l
        Send, {Numpad1}
        Send, {Numpad1}
        break
    }
    Return

;==================測試功能模組1==================  
^+I::
    Gosub, slip
    ToolTip, 判斷銷貨單為新增狀態中..., 900, 300
    loop {
        ToolTip, 等待銷單為新增狀態中...., 900, 300
        ControlGet, OutputVar, Visible,, Edit86, ahk_class ThunderRT6MDIForm
        if (OutputVar = 0) {
            Sleep % 100
            Send,{F2}
            break
        }
        else {
            loop {
                ControlGetText, OutputVar, Edit36, ahk_class ThunderRT6MDIForm
                if StrLen(OutputVar) = 0 {
                break
                }
                else {
                    ToolTip, 等待銷單為退出狀態中...., 900, 300
                    Gosub, main2
                    Gosub, confirm
                    Gosub, slip
                    ToolTip, 等待銷單為新增狀態中...., 900, 300
                    loop {
                        ControlGet, OutputVar, Visible,, Edit86, ahk_class ThunderRT6MDIForm
                        if (OutputVar = 0) {
                            Sleep % 100
                            Send,{F2}
                            break
                        }
                    }
                }
            }
        }
        break
    }

    ToolTip, 銷單新增完成, 900, 300

    Return

;==================測試功能模組2================== 
^+O::
    ToolTip, 打開銷單視窗中..., 900, 300
    WinActivate ahk_class ThunderRT6MDIForm
    WinWait, ahk_class ThunderRT6MDIForm
    loop {
        WinGetText,Str,A
        Haystack := Str
        Needle := "銷貨單"
        Atr := InStr(Haystack, Needle)
        if (Atr = 1) {
            ToolTip, 等待銷單為退出狀態中...., 900, 300
            Loop {
                ControlGet, OutputVar, Visible,, Edit86, ahk_class ThunderRT6MDIForm
                if (OutputVar = 1) {
                    Sleep % 100
                    Send,{Esc}
                    Sleep % 500
                    Gosub, confirm
                    break
                }
                else {
                    ToolTip, 銷貨單等待目標視窗錯誤，請手動切換或是重新流程！, 900, 300
                    break
                }
            }
            break
        }
        else {
            ToolTip, 重新開啟銷單畫面中....., 900, 300
            WinGetText,Str,A
            Haystack := Str
            Needle := "功能快捷視窗"
            Atr := InStr(Haystack, Needle)
            while not (Atr = 1) {
                Sleep % 100
                Send, {F12}
                Sleep % 100
                break
            }
            WinGetText,Str,A
            Haystack := Str
            Needle := "銷貨單"
            Atr1 := InStr(Haystack, Needle)
            while not (Atr1 = 1) {
                Sleep % 100
                Send, {F10}l11
                break
            }
            continue
        }
    }

    ToolTip, 開啟銷單完成, 900, 300
    
    Return