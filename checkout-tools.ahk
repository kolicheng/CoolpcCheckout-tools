;==================開發票結帳功能模組==================
Pause::
    Gui, Destroy
    Gosub Finalcheck

    __title := "開發票-實收金額"
    __text := ""
    InputBox, inputE,%__title%,%__text%,,200,100
    if (ErrorLevel = 1) {
        Gosub, stop
    }
    ControlGetText, OutputVar , Edit84, ahk_class ThunderRT6MDIForm, 銷貨單
    A := inputE
    B := StrReplace(OutputVar, "," )
    Var := A - B
    MsgBox, 260, 結帳小工具, 需要電子發票載具嗎？
    IfMsgBox Yes
    {
        __title := "請輸入載具號碼"
        __text := ""
        InputBox, Haystack,%__title%,%__text%,,200,100
        if (ErrorLevel = 1) {
            Gosub, stop
        }
        inputA := Haystack
        Needle := "/"
        If InStr( inputA, Needle){
            inputA := Haystack
        }
        else {
            Gosub, stop
        }
    }
    else {
        inputA := ""
    }

    __title := "請輸入統一編號"
    __text := ""
    InputBox, __invoice,%__title%,%__text%,,200,100
    if (ErrorLevel = 1) {
        Gosub, stop
    }
    inputB := StrLen(__invoice)
    if (inputB = 8 or inputB = 0) {
        inputB := __invoice
    }
    else {
        Gosub, stop
    }
    MsgBox, 260, 結帳小工具, 刷卡嗎？
    IfMsgBox Yes
    {
        __title := "請輸入卡號後4碼"
        __text := ""
        InputBox, inputC,%__title%,%__text%,,200,100
        ControlFocus, Edit23, ahk_class ThunderRT6MDIForm
        Control, EditPaste, 刷卡/%inputC%%A_Space%, Edit30, ahk_class ThunderRT6MDIForm
        if (ErrorLevel = 1) {
            Gosub, stop
        }
    }
    else {
        inputC := ""
    }

    Gosub, run1

    WinActivate ahk_class ThunderRT6MDIForm
    Control, Check,, Button5, ahk_class ThunderRT6MDIForm, 銷貨單
    ToolTip, 發票視窗開啟中..., 900, 300
    WinWait, ahk_class ThunderRT6FormDC, 重載貨單明細
    WinGetText, OutputVar , ahk_class ThunderRT6FormDC, 重載貨單明細
    control = %OutputVar%
    while (control := "重載貨單明細") {
        ToolTip, 等待發票號碼產生中..., 900, 300
        SetControlDelay 0
        ControlClick, Edit5, ahk_class ThunderRT6FormDC,,,, NA
        loop {
            ControlGetText, OutputVar, Edit5, ahk_class ThunderRT6FormDC
            if StrLen(OutputVar) = 0 {
                Sleep, 1000
                continue
            }
            else {
                ControlGetText, OutputVarEdit48, Edit48, ahk_class ThunderRT6FormDC
                if StrLen(OutputVarEdit48) = 0 {
                    Control, EditPaste, %inputA%, Edit48, ahk_class ThunderRT6FormDC
                }
                ControlGetText, OutputVaredit2, Edit2, ahk_class ThunderRT6FormDC
                if StrLen(OutputVaredit2) = 0 {
                    Control, EditPaste, %inputB%, Edit2, ahk_class ThunderRT6FormDC
                }
                ControlGetText, OutputVarEdit14, Edit14, ahk_class ThunderRT6FormDC
                if (OutputVarEdit14 = "") {
                    Control, EditPaste, %inputC%%A_Space%, Edit14, ahk_class ThunderRT6FormDC
                }
                break
            }
        }
        break
    }

    Sleep % 100
    Send,{F9}
    Send, {Esc}
    Gosub, slip
    Send,{F9}

    Gosub, Print
    Gosub, slip

    ControlFocus, Edit52, ahk_class ThunderRT6MDIForm
    send, {Enter}

    Gosub, stoptip
    Return

;==================帶入客訂單功能模組==================
Insert::
    Gui, Destroy

    __title := "帶單-後4碼/完整8碼"
    __text := ""
    OutputVar := ""
    InputBox, __textFormatTime, %__title%, %__text%,, 200, 100,,,,, %OutputVar%
    if (ErrorLevel = 1) {
        Gosub, stop
    }
    inputI := StrLen(__textFormatTime)
    if (inputI = 4) {
        FormatTime, OutputVar,,yyyyMMdd
        inputH = %OutputVar%%__textFormatTime%
    }
    else if (inputI = 12) {
        inputH = %__textFormatTime%
    }
    else {
        Gosub, stop
    }
    Gosub, run1
    Gosub, gosales1
    Gosub, main1
    WinWait, ahk_class ThunderRT6MDIForm, 銷貨單
    Loop {
        Sleep % 100
        WinGetText, OutputVar , ahk_class ThunderRT6MDIForm, 銷貨單
        control = %OutputVar%
        if (control != 銷貨單) {
            Sleep % 100
            ControlFocus, Edit36, ahk_class ThunderRT6MDIForm
            Sleep % 100
            ControlSetText , Edit36,, ahk_class ThunderRT6MDIForm
            Sleep % 100
            Control, EditPaste, %inputH%, Edit36, ahk_class ThunderRT6MDIForm
            Sleep % 100
            ControlSend, Edit36, {Enter}, ahk_class ThunderRT6MDIForm
            break
        }
    }
    WinWait, ahk_class ThunderRT6FormDC, 區域選取
    Loop {
        Sleep % 100
        ControlGetFocus, focused_control, ahk_class ThunderRT6FormDC, 區域選取
        control = %focused_control%
        if (control != TG70.ApexGridOleDB32.202) {
            Sleep % 100
            Send, {F9}
            break
        }
    }

    Gosub, stoptip

    Return

;==================複製銷單功能模組==================
ScrollLock::
    __title := "列印-輸入銷單後四碼"
    __text := ""
    OutputVar := ""
    InputBox, __textFormatTime, %__title%, %__text%,, 200, 100,,,,, %OutputVar%
    if (ErrorLevel = 1) {
        Gosub, stop
    }
    inputT := StrLen(__textFormatTime)
    if (inputT = 4) {
        FormatTime, OutputVar,,yyyyMMdd
        inputG= %OutputVar%%__textFormatTime%
    }
    else if (inputT = 12) {
        inputG= %__textFormatTime%
    }
    else {
        Gosub, stop
    }

    Gosub, run1
    Gosub, gosales
    Gosub, find

    WinWait, ahk_class ThunderRT6FormDC, 軟體類別
    Loop {
        Sleep % 100
        WinGetText, OutputVar , ahk_class ThunderRT6FormDC, 軟體類別
        control = %OutputVar%
        if (control != 軟體類別) {
            Sleep % 100
            ControlSetText , Edit22, %    /  /  , ahk_class ThunderRT6FormDC
            Sleep % 100
            Control, EditPaste, %inputG%, Edit24, ahk_class ThunderRT6FormDC
            Sleep % 100
            Control, EditPaste, %inputG%, Edit25, ahk_class ThunderRT6FormDC
            Sleep % 100
            Send, {F9}
            break
        }
        else {
            Sleep % 100
        }
    
    }
    Gosub, slip

    ControlFocus, Edit23, ahk_class ThunderRT6MDIForm
    ToolTip, 銷貨單視窗確認中..., 900, 300
    loop {
        WinGetText,Str,A
        Haystack := Str
        Needle := "銷貨單"
        Atr := InStr(Haystack, Needle)
        if (Atr = 1) {
            Sleep % 100
            Send,{F7}
            Sleep % 100
            break
        }
        else{
            ToolTip, 等待銷單視窗開啟中...., 900, 300
            break
        }
    }
    
    Gosub, Print1
    Gosub, stoptip

    Return

;==================拷貝銷單功能模組==================
^+C::
    Gui, Destroy

    __title := "拷貝-銷單後4碼或完整8碼"
    __text := ""
    OutputVar := ""
    InputBox, __textFormatTime, %__title%, %__text%,, 200, 100,,,,, %OutputVar%
    if (ErrorLevel = 1) {
        Gosub, stop
    }
    inputT := StrLen(__textFormatTime)
    if (inputT = 4) {
        FormatTime, OutputVar,,yyyyMMdd
        inputR = %OutputVar%%__textFormatTime%
    }
    else if (inputT = 12) {
        inputR = %__textFormatTime%
    }
    else {
        Gosub, stop
    }

    Gosub, run1

    WinWait, ahk_class ThunderRT6MDIForm, 銷貨單

    Gosub, gosales
    Gosub, find

    WinWait, ahk_class ThunderRT6FormDC, 軟體類別
    Loop {
        Sleep % 100
        WinGetText, OutputVar , ahk_class ThunderRT6FormDC, 軟體類別
        control = %OutputVar%
        if (control != 軟體類別) {
            Sleep % 100
            ControlSetText , Edit22, %    /  /  , ahk_class ThunderRT6FormDC
            Sleep % 100
            Control, EditPaste, %inputR%, Edit24, ahk_class ThunderRT6FormDC
            Sleep % 100
            Control, EditPaste, %inputR%, Edit25, ahk_class ThunderRT6FormDC
            Sleep % 100
            Send, {F9}
            break
        }
        else {
            Sleep % 100
        }
    }

    ControlFocus, Edit23, ahk_class ThunderRT6MDIForm

    Gosub, slip
    Gosub, main1

    Loop {
        Sleep % 100
        ControlGet, OutputVar, Visible,, Edit86, ahk_class ThunderRT6MDIForm, 銷貨單
        control = %OutputVar%
        if (control = 1) {
            Sleep % 100
            Send, !C
            break
        }
    }

    WinWait, ahk_class ThunderRT6FormDC

    Loop {
        Sleep % 100
        WinGetText, OutputVar , ahk_class ThunderRT6FormDC, 產生單據日期
        control = %OutputVar%
        if (control != 產生單據日期) {
            Sleep % 100
            SetControlDelay -1
            ControlClick, Edit5, ahk_class ThunderRT6FormDC,,,, NA
            Sleep % 100
            Control, EditPaste, %inputR%, Edit3, ahk_class ThunderRT6FormDC
            Sleep % 100
            Control, Check,, Button1, ahk_class ThunderRT6FormDC
            break
        }
    }

    WinWait, ahk_class ThunderRT6FormDC

    Loop {
        Sleep % 100
        PixelGetColor, color, 468, 331 , Alt
        control = %color%
        if (control = 0xD77800) {
            Sleep % 100
            CoordMode, Mouse , Screen
            __ClickX:=524
            __ClickY:=458
            __ClickTimes:=1
            Click %__ClickX%, %__ClickY%, %__ClickTimes%
            Sleep % 100
            Send, {F9}
            Sleep % 100
            break
        }
    }

    WinWait, ahk_class ThunderRT6FormDC, 區域選取

    Loop {
        Sleep % 100
        ControlGetFocus, focused_control, ahk_class ThunderRT6FormDC, 區域選取
        control = %focused_control%
        if (control != TG70.ApexGridOleDB32.202) {
            Sleep % 100
            Send, {F9}
            break
        }
    }

    WinWait, ahk_class ThunderRT6MDIForm, 銷貨單

    Loop {
        WinGetText, OutputVar , ahk_class ThunderRT6MDIForm, 銷貨單
        control = %OutputVar%
        if (control != 銷貨單) {
            ControlGetText, OutputVarmark, Edit39, ahk_class ThunderRT6MDIForm
            ControlGetText, OutputVarsales1, Edit22, ahk_class ThunderRT6MDIForm
            ControlGetText, OutputVarsales2, Edit23, ahk_class ThunderRT6MDIForm
            EnvSet, mark, %OutputVarmark%
            EnvSet, sales1, %OutputVarsales1%
            EnvSet, sales2, %OutputVarsales2%
            Gosub, markname
            Sleep % 1000
        }
        break
    }

    WinWait, ahk_class ThunderRT6MDIForm, 銷貨單

    Loop {
        WinGetText, OutputVar , ahk_class ThunderRT6MDIForm, 銷貨單
        control = %OutputVar%
        if (control != 銷貨單) {
            Sleep % 1000
            Gosub, Salesname
        }
        break
    }

    WinWait, ahk_class ThunderRT6MDIForm, 銷貨單

    Loop {
        Sleep % 100
        WinGetText, OutputVar , ahk_class ThunderRT6MDIForm, 銷貨單
        control = %OutputVar%
        if (control != 銷貨單) {
            Sleep % 100
            Control, EditPaste, 原銷單%inputR%%A_Space%, Edit30, ahk_class ThunderRT6MDIForm, 銷貨單
            break
        }
    }

    Gosub, stoptip

    Return

;==================文字快捷鍵1模組功能==================
^E::
    Send, 你好  請問可以跟你們調一個嗎? 謝謝
    return

;==================文字快捷鍵2模組功能==================
^W::
    Send, 中租分期  姓名: 仲信案編: 期數:期
    return

;==================快捷鍵說明模組功能==================
^+H::
    CoordMode, ToolTip
    ToolTip, 結帳：        Psuse`r列印發票：ScrollLock`r帶客訂單：Insert`r複製銷單：Ctrl+Shift+C`r緊急停止：F8`r`r提示介面五秒自動消失
    Sleep % 10000
    ToolTip
    Return

;==================緊急停止功能模組==================
F8::
    Gui, Destroy
    MsgBox, 4624, 結帳小工具, 手動取消
    reload
    Return

;==================OSD GUI模組==================
Finalcheck:
    CustomColor := "Default"
    Gui, +AlwaysOnTop +ToolWindow
    Gui, Color, %CustomColor%
    Gui, Font, s14 w600, Arial
    Gui, Add, Text, v1text, XXXXX YYYYY
    Gui, Font
    Gui, Font, s14 w600, Arial
    Gui, Add, Text, v2text, XXXXX YYYYY
    Gui, Font
    Gui, Font, s14 w600 c0xFF0000, Arial
    Gui, Add, Text, v3text, XXXXX YYYYY
    Gui, Font
    Gui, Font, s14 w600, Arial
    Gui, Add, Text, v4text, XXXXX YYYYY
    Gui, Font
    Gui, Font, s14 w600 c0xFF0000, Arial
    Gui, Add, Text, v5text, XXXXX YYYYY
    Gui, Font
    Gui, Font, s14 w600, Arial
    Gui, Add, Text, v6text, XXXXX YYYYY
    Gui, Font
    Gui, Add, Text, x20 y15, 實收金額：
    Gui, Add, Text, x20 y45, 開立金額：
    Gui, Add, Text, x20 y75, 找零金額：
    Gui, Add, Text, x190 y15, 載具號碼：
    Gui, Add, Text, x190 y45, 統一編號：
    Gui, Add, Text, x190 y75, 卡號四碼：
    WinSet, TransColor, %CustomColor% 100
    SetTimer, UpdateOSD, 200
    Gosub, UpdateOSD
    Gui, Show, x670 y10 w370 h100 NoActivate, 確認視窗
    return

;==================OSD模組==================
UpdateOSD:
    GuiControl,, 1text, %inputE%
    GuiControl,, 2text, %B%
    GuiControl,, 3text, %Var%
    GuiControl,, 4text, %inputA%
    GuiControl,, 5text, %inputB%
    GuiControl,, 6text, %inputC%
    GuiControl, Move, 1text, x80 y10 w100 h20
    GuiControl, Move, 2text, x80 y39 w100 h20
    GuiControl, Move, 3text, x80 y69 w100 h20
    GuiControl, Move, 4text, x250 y10 w100 h20
    GuiControl, Move, 5text, x250 y39 w100 h20
    GuiControl, Move, 6text, x250 y69 w100 h20
    return

;==================載入全域紀錄模組==================
Salesname:
    EnvGet, OutputVar2, sales1
    EnvGet, OutputVar3, sales2
    loop {
        ControlGetFocus, OutputVar , ahk_class ThunderRT6MDIForm
        if (OutputVar := "Edit27") {
            ControlSetText , Edit22, %OutputVar2%, ahk_class ThunderRT6MDIForm
            ControlSetText , Edit23, %OutputVar3%, ahk_class ThunderRT6MDIForm
            break
        }
        else {
            Sleep % 100
        }
    }
    EnvSet, sales1
    EnvSet, sales2
    Return

;==================匯出全域紀錄模組==================
Markname:
    EnvGet, OutputVar1, mark

    Needlep1 := "w501"
    Needlep2 := "W502"
    Needlek1 := "5X01"
    Needlem1 := "V5001"

    if InStr(OutputVar1, Needlep1) {
        ControlClick, Edit39, ahk_class ThunderRT6MDIForm,,,, NA
        ControlSetText, Edit39,, ahk_class ThunderRT6MDIForm
        ControlSendRaw, Edit39, 308`n, ahk_class ThunderRT6MDIForm
    }
    else if InStr(OutputVar1, Needlep2) {
        ControlClick, Edit39, ahk_class ThunderRT6MDIForm,,,, NA
        ControlSetText, Edit39,, ahk_class ThunderRT6MDIForm
        ControlSendRaw, Edit39, 312`n, ahk_class ThunderRT6MDIForm
    }
    else if InStr(OutputVar1, Needlek1) {
        ControlClick, Edit39, ahk_class ThunderRT6MDIForm,,,, NA
        ControlSetText, Edit39,, ahk_class ThunderRT6MDIForm
        ControlSendRaw, Edit39, 352`n, ahk_class ThunderRT6MDIForm
    }
    else if InStr(OutputVar1, Needlem1) {
        ControlClick, Edit39, ahk_class ThunderRT6MDIForm,,,, NA
        ControlSetText, Edit39,, ahk_class ThunderRT6MDIForm
        ControlSendRaw, Edit39, 394`n, ahk_class ThunderRT6MDIForm
    }

    EnvSet, mark

    Return

;==================打開銷貨單模組（打開其他視窗情況）==================
gosales1:
    ToolTip, 確認銷單視窗中..., 900, 300
    WinActivate ahk_class ThunderRT6MDIForm
    WinWait, ahk_class ThunderRT6MDIForm
    loop {
        WinGetText,Str,A
        Haystack := Str
        Needle := "銷貨單"
        Atr := InStr(Haystack, Needle)
        if (Atr = 1) {
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
        }
    }

    ToolTip, 開啟銷單完成, 900, 300
    Return

;==================打開銷貨單模組（未打開情況）==================
gosales:
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

;==================停止提示視窗功能==================
stop:
    ToolTip
    Gui, Destroy
    MsgBox, 262704, 結帳小工具, 錯誤，取消動作！
    reload
    Return

;==================檢查銷貨單狀態功能==================
main1:
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

;==================檢查銷貨單狀態功能（其他視窗）==================
main2:
    Gosub, slip
    ToolTip, 判斷銷貨單為退出狀態中..., 900, 300
    Loop {
        ControlGet, OutputVar, Visible,, Edit86, ahk_class ThunderRT6MDIForm
        if (OutputVar = 1) {
            loop {
                WinGetText,Str,A
                Haystack := Str
                Needle := "銷貨單"
                Atr := InStr(Haystack, Needle)
                if (Atr = 1) {
                    Sleep % 100
                    Send,{Esc}
                    Sleep % 500
                    break
                }
                else{
                    Sleep % 100
                    ToolTip, 等待銷單為退出狀態中...., 900, 300
                    break
                }
            }
        }
        break
    }

    ToolTip, 銷單退出完成, 900, 300
    
    Return

;==================提示視窗功能（執行中）==================
run1:
    CoordMode, ToolTip
    ToolTip, 小工具執行中....., 900, 300
    Return

;==================提示視窗功能（完成）==================
stoptip:
    CoordMode, ToolTip
    ToolTip, 動作完成!!!!, 900, 300
    Sleep % 5000
    ToolTip
    Return

;==================判斷銷單狀態==================
slip:
    ToolTip, 銷貨單視窗確認中..., 900, 300
    loop {
        WinGet, ActiveWindowID, ID, A
        WinGetText, VisibleText, ahk_id %ActiveWindowID%
        if (SubStr(VisibleText, 1, 3) = "銷貨單")
    {
    break
    }
    else
    {
    ToolTip, 等待銷單視窗開啟中...., 900, 300
    Sleep % 100
    }
    }
    Return

;==================判斷銷單狀態（判斷詳查視窗）==================
find:
    Gosub, slip
    ControlFocus, Edit23, ahk_class ThunderRT6MDIForm
    ToolTip, 詳查視窗載入中..., 900, 300
    WinGetText,Str,A
    Haystack := Str
    Needle := " 審核確認"
    Atr := InStr(Haystack, Needle)
    while not (Atr = 1) {
        Sleep % 100
        Send,{F5}
        Sleep % 100
        break
    }

    Return

;==================判斷銷單狀態（發票視窗）==================
receipt:
    ToolTip, 發票視窗開啟中..., 900, 300
    loop {
        WinGetText,Str,A
        Haystack := Str
        Needle := "多選 "
        Atr := InStr(Haystack, Needle)
        if (Atr = 1) {
            break
        }
        else{
            ToolTip, 等待發票視窗開啟中....., 900, 300
            break
        }
    }
    Return

;==================判斷銷單狀態（載入中）==================
load:
    ToolTip, 客訂單載入中..., 900, 300
    loop {
        WinGetText,Str,A
        Haystack := Str
        Needle := "▼"
        Atr := InStr(Haystack, Needle)
        if (Atr = 1) {
            break
        }   
        else{
            ToolTip, 等待客訂單載入畫面中...., 900, 300
            break
        }
    }
    Return

;==================判斷銷單狀態（確認中）==================
confirm:
    ToolTip, 等待確認視窗中..., 900, 300
    loop {
        WinGetText,Str,A
        Haystack := Str
        Needle := "確定"
        Atr := InStr(Haystack, Needle)
        if (Atr = 1) {
            Sleep % 100
            SetControlDelay -1
            ControlClick, ThunderRT6CommandButton1, ahk_class ThunderRT6FormDC ,確定,,, NA
        break
        }
        else{
            Sleep % 100
            continue
        }
    }

    ToolTip, 確認流程完成, 900, 300

    Return

;==================判斷銷單狀態（列印中）==================
Print:
    ToolTip, 列印發票中..., 900, 300

    Loop {
        ControlGet, OutputVar, Visible,, ThunderRT6PictureBoxDC1, A
        if (OutputVar = 1) {
            Send,{Esc}
            break
        }
        else {
            Sleep % 100
        }
    }

    Loop {
        ControlGet, OutputVar, Visible,, ThunderRT6CommandButton6, A
        if (OutputVar = 1) {
            SetControlDelay -1
            ControlClick, ThunderRT6CommandButton6, ahk_class ThunderRT6FormDC,,,, NA
            break
        }
        else {
            Sleep % 100
        }
    }

    ToolTip, 發票流程完成, 900, 300
    Return

;==================判斷銷單狀態（列印發票中）==================
Print1:
    ToolTip, 列印發票中..., 900, 300
    Loop {
        ControlGet, OutputVar, Visible,, ThunderRT6CommandButton6, A
        if (OutputVar = 1) {
            SetControlDelay -1
            ControlClick, ThunderRT6CommandButton6, ahk_class ThunderRT6FormDC,,,, NA
            break
        }
        else {
            Sleep % 100
        }
    }

    Loop {
        ControlGet, OutputVar, Visible,, ThunderRT6PictureBoxDC1, A
        if (OutputVar = 1) {
            Send,{Esc}
            break
        }
        else {
            Sleep % 100
        }
    }

    Loop {
        ControlGet, OutputVar, Visible,, ThunderRT6CommandButton1, A
        if (OutputVar = 1) {
            SetControlDelay -1
            ControlClick, ThunderRT6CommandButton1, ahk_class ThunderRT6FormDC,,,, NA
            break
        }
        else {
            Sleep % 100
        }
    }

    ToolTip, 發票流程完成, 900, 300

    Return

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