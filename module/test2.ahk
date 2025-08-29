#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

; 主要腳本啟動的地方
GoGui:
    ; 如果視窗已經存在就關閉它，以便重新整理
    IfWinExist, 文字快捷選單
        Gui, Destroy

    ; 建立GUI選單
    Gui, Add, Text, , 請選擇要輸入的文字：
    
    ; 讀取文字檔並動態建立按鈕
    y_pos := 30
    Loop, read, text_options.txt
    {
        Gui, Add, Button, x10 y%y_pos% w200 h30 gGoText, %A_LoopReadLine%
        y_pos += 40
    }
    
    ; 加入重新整理和關閉按鈕
    Gui, Add, Button, x10 y%y_pos% w95 h30 gGoGui, 重新整理
    Gui, Add, Button, x115 y%y_pos% w95 h30 gGuiClose, 關閉
    
    Gui, Show, , 文字快捷選單
Return

; GUI關閉時的動作
GuiClose:
ExitApp

; 這個是核心功能，所有按鈕都會跑這裡
GoText:
    ; 記住你點擊按鈕前的視窗
    LastActiveWinID := WinExist("A")
    
    ; 取得你點擊的按鈕上的文字
    ButtonText := A_GuiControl
    
    Gui, Hide ; 隱藏選單視窗
    
    WinActivate, ahk_id %LastActiveWinID% ; 切換回原本的視窗
    Sleep, 100 ; 稍微等一下
    
    Send, %ButtonText% ; 輸入文字
    return

; 隱藏GUI的快捷鍵
^!s::
    ; 檢查檔案是否存在
    IfNotExist, text_options.txt
    {
        ; 如果檔案不存在，跳出訊息視窗
        MsgBox, 0, 提醒, 偵測不到「text_options.txt」檔案。%A_Space%
        (LTrim
            現在將為您建立一個新的檔案，
            請在記事本中輸入您要的文字，
            每行一個，完成後儲存即可。
        )
        ; 自動打開記事本並帶入檔名
        Run, notepad.exe "text_options.txt"
        Return ; 結束
    }
    
    ; 如果檔案存在，就顯示GUI選單
    GoSub, GoGui
Return