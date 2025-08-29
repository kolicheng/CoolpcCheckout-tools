; 使用獨立視窗顯示快捷鍵說明
^+H::
    ; 檢查視窗是否已經存在，如果存在就顯示
    IfWinExist, 快捷鍵說明
        Gui, Show
    Else
    {
        ; 如果不存在，就建立一個新視窗
        Gui, Add, Text,, 結帳：Psuse`r列印發票：ScrollLock`r帶客訂單：Insert`r複製銷單：Ctrl+Shift+C`r緊急停止：F8
        Gui, Add, Button, gGuiClose, 關閉
        Gui, Show, , 快捷鍵說明
    }
    Return

; 關閉視窗的指令
GuiClose:
    Gui, Destroy
    Return