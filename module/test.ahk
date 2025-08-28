; 設定腳本持續運作並偵測滑鼠
#Persistent
#InstallMouseHook

; 設定計時器，每100毫秒執行一次
SetTimer, CheckVisibleText, 100

; 定義函式
CheckVisibleText:
    ; 取得滑鼠下方的視窗ID
    MouseGetPos, , , MouseWindowID

    ; 取得該視窗中所有的可見文字
    WinGetText, VisibleText, ahk_id %MouseWindowID%

    ; 檢查這些文字的前三個字是不是 "銷貨單"
    if (SubStr(VisibleText, 1, 3) = "銷貨單")
    {
        ; 如果是，顯示 "是"
        ToolTip, 是
    }
    else
    {
        ; 如果不是，顯示 "否"
        ToolTip, 否
    }
Return

; 按下 Esc 鍵來結束腳本
Esc::
    ExitApp