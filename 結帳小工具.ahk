;====================熱鍵定義區塊====================
; 先設定預設值
Hotkey_結帳 = Pause
Hotkey_帶入客訂單 = Insert
Hotkey_直接列印發票 = ^+C
Hotkey_複製銷單 = ScrollLock
Hotkey_緊急停止 = F8
Hotkey_快速輸入 = ^E
Hotkey_快捷鍵說明 = ^+H
Hotkey_修改熱鍵 = ^+G
Hotkey_全域設定 = ^+S

; 設定 OSD 預設值
OSD_enabled := 1 ; 預設開啟 OSD

; 檢查 Hotkeys.ini 檔案是否存在
if FileExist("Hotkeys.ini") {
    ; 如果檔案存在，讀取熱鍵設定
    IniRead, Hotkey_結帳, Hotkeys.ini, Hotkeys, 結帳
    IniRead, Hotkey_帶入客訂單, Hotkeys.ini, Hotkeys, 帶入客訂單
    IniRead, Hotkey_直接列印發票, Hotkeys.ini, Hotkeys, 直接列印發票
    IniRead, Hotkey_複製銷單, Hotkeys.ini, Hotkeys, 複製銷單
    IniRead, Hotkey_緊急停止, Hotkeys.ini, Hotkeys, 緊急停止
    IniRead, Hotkey_快速輸入, Hotkeys.ini, Hotkeys, 快速輸入
    IniRead, Hotkey_快捷鍵說明, Hotkeys.ini, Hotkeys, 快捷鍵說明
    IniRead, Hotkey_修改熱鍵, Hotkeys.ini, Hotkeys, 修改熱鍵
    IniRead, Hotkey_全域設定, Hotkeys.ini, Hotkeys, 全域設定

    ; 檢查 OSD 設定是否存在，如果沒有，就用預設值 1
    IniRead, OSD_enabled, Hotkeys.ini, Settings, OSD
    if (ErrorLevel) {
        OSD_enabled := 1
    }

    ; 檢查「已執行過」的變數是否存在
    IniRead, run_status, Hotkeys.ini, Settings, RunStatus
    if (run_status = "1") {
        ; 如果找到 RunStatus=1，表示不是第一次啟動，什麼都不做
    } else {
        ; 如果沒有找到或值不為1，表示需要重新顯示說明
        Gosub, Label_快捷鍵說明
        IniWrite, 1, Hotkeys.ini, Settings, RunStatus
    }
} else {
    ; 如果檔案不存在，表示是第一次啟動
    ; 顯示快捷鍵說明
    Gosub, Label_快捷鍵說明
    
    ; 將預設熱鍵存入新的檔案
    IniWrite, %Hotkey_結帳%, Hotkeys.ini, Hotkeys, 結帳
    IniWrite, %Hotkey_帶入客訂單%, Hotkeys.ini, Hotkeys, 帶入客訂單
    IniWrite, %Hotkey_直接列印發票%, Hotkeys.ini, Hotkeys, 直接列印發票
    IniWrite, %Hotkey_複製銷單%, Hotkeys.ini, Hotkeys, 複製銷單
    IniWrite, %Hotkey_緊急停止%, Hotkeys.ini, Hotkeys, 緊急停止
    IniWrite, %Hotkey_快速輸入%, Hotkeys.ini, Hotkeys, 快速輸入
    IniWrite, %Hotkey_快捷鍵說明%, Hotkeys.ini, Hotkeys, 快捷鍵說明
    IniWrite, %Hotkey_修改熱鍵%, Hotkeys.ini, Hotkeys, 修改熱鍵
    IniWrite, %Hotkey_全域設定%, Hotkeys.ini, Hotkeys, 全域設定

    ; 儲存 OSD 預設值到檔案
    IniWrite, %OSD_enabled%, Hotkeys.ini, Settings, OSD

    ; 存入「已執行過」的變數
    IniWrite, 1, Hotkeys.ini, Settings, RunStatus
	FileSetAttrib, +H, Hotkeys.ini
}

;====================啟動熱鍵區塊====================

Hotkey, %Hotkey_結帳%, Label_結帳
Hotkey, %Hotkey_帶入客訂單%, Label_帶入客訂單
Hotkey, %Hotkey_直接列印發票%, Label_直接列印發票
Hotkey, %Hotkey_複製銷單%, Label_複製銷單
Hotkey, %Hotkey_緊急停止%, Label_緊急停止
Hotkey, %Hotkey_快速輸入%, Label_快速輸入
Hotkey, %Hotkey_快捷鍵說明%, Label_快捷鍵說明
Hotkey, %Hotkey_修改熱鍵%, Label_修改熱鍵
Hotkey, %Hotkey_全域設定%, Label_全域設定

return

;==================全局變數區塊==================
; 設定一個「開關」，用於防止多個功能同時執行。
is_running_flag := 0

;==================帶入客訂單功能模組==================
Label_帶入客訂單:
	; 檢查開關，如果正在執行中，就立即停止。
	if is_running_flag {
		Return
	}
	; 打開開關，表示功能正在執行。
	is_running_flag := 1

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
	
	WinWait, ahk_class ThunderRT6MDIForm, 銷貨單

	Loop {
		WinGetText, OutputVar , ahk_class ThunderRT6MDIForm, 銷貨單
		control = %OutputVar%
		if (control != 銷貨單) {
			ControlGetText, OutputVarmark, Edit38, ahk_class ThunderRT6MDIForm
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

	Gosub, stoptip
	
	; 關閉開關，表示功能執行完畢。
	is_running_flag := 0
	Return

;==================開發票結帳功能模組==================
Label_結帳:
	; 檢查開關，如果正在執行中，就立即停止。
	if is_running_flag {
		Return
	}
	; 打開開關，表示功能正在執行。
	is_running_flag := 1
	
	; 如果 OSD_enabled 為 1，則執行 Finalcheck 函式
	if (OSD_enabled = 1) {
		Gosub, Finalcheck
	} else {
		Gui, Destroy
	}

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
	MsgBox, 260, "結帳小工具, 需要電子發票載具嗎？"
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
	Send, {Enter}

	Gosub, stoptip

	; 關閉 OSD 視窗並停止更新
	Gui, Destroy
	SetTimer, UpdateOSD, Off

	; 關閉開關，表示功能執行完畢。
	is_running_flag := 0
	Return

;==================直接列印發票功能模組==================
Label_直接列印發票:
	; 檢查開關，如果正在執行中，就立即停止。
	if is_running_flag {
		Return
	}
	; 打開開關，表示功能正在執行。
	is_running_flag := 1
	
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
	
	; 關閉開關，表示功能執行完畢。
	is_running_flag := 0
	Return

;==================拷貝銷單功能模組==================
Label_複製銷單:
	; 檢查開關，如果正在執行中，就立即停止。
	if is_running_flag {
		Return
	}
	; 打開開關，表示功能正在執行。
	is_running_flag := 1
	
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
		ControlGet, OutputVar, Visible,, Edit91, ahk_class ThunderRT6MDIForm, 銷貨單
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
			ControlGetText, OutputVarmark, Edit38, ahk_class ThunderRT6MDIForm
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

	; 關閉開關，表示功能執行完畢。
	is_running_flag := 0
	Return

;==================文字輸入快捷鍵模組功能==================
Label_快速輸入:
	; 檢查開關，如果正在執行中，就立即停止。
	if is_running_flag {
		Return
	}
	; 打開開關，表示功能正在執行。
	is_running_flag := 1
	
	; 檢查檔案是否存在
	IfNotExist, text_options.txt
	{
		; 這裡存放你要自動加入的文字
		DefaultText =
		(LTrim
			你好 請問可以跟你們調一個嗎? 謝謝
			中租分期 姓名: 仲信案編: 期數:期
		)
		; 設定接下來的檔案操作都使用 UTF-8 編碼
		FileEncoding, UTF-8
		; 建立檔案並寫入範例文字
		FileAppend, %DefaultText%, text_options.txt
		; 跳出訊息視窗
		MsgBox, 0, 提醒, 偵測不到「text_options.txt」檔案。%A_Space%
		(LTrim
			現在將為您建立一個新的檔案！已自動填入範例文字，請直接儲存即可使用。
		)
		; 自動打開記事本並帶入檔名
		Run, notepad.exe "text_options.txt"
		Return ; 結束
	}
	; 如果檔案存在，就顯示GUI選單
	GoSub, Label_GoTextGui

	; 關閉開關，表示功能執行完畢。
	is_running_flag := 0
	Return

Label_GoTextGui:
	; 如果視窗已經存在就關閉它，以便重新整理
	IfWinExist, 文字快捷選單
		Gui, Destroy

	; 建立GUI選單
	Gui, Add, Text, , 請選擇要輸入的文字：

	; 讀取文字檔並動態建立按鈕
	FileEncoding, UTF-8  ; 讀取檔案前也需要指定編碼，確保正確
	y_pos := 30
	Loop, read, text_options.txt
	{
		Gui, Add, Button, x10 y%y_pos% w200 h30 gLabel_GoText, %A_LoopReadLine%
		y_pos += 40
	}

	; 加入重新整理、編輯、新增和關閉按鈕
	Gui, Add, Button, x10 y%y_pos% w60 h30 gLabel_GoAdd, 新增
	Gui, Add, Button, x80 y%y_pos% w60 h30 gLabel_GoEdit, 編輯
	Gui, Add, Button, x150 y%y_pos% w60 h30 gLabel_GoTextGui, 重新整理
	Gui, Add, Button, x220 y%y_pos% w60 h30 gLabel_GuiClose, 關閉
	Gui, Show, , 文字快捷選單
	Return

Label_GoText:
	; 記住你點擊按鈕前的視窗
	LastActiveWinID := WinExist("A")

	; 取得你點擊的按鈕上的文字
	ButtonText := A_GuiControl
	Gui, Hide ; 隱藏選單視窗
	WinActivate, ahk_id %LastActiveWinID% ; 切換回原本的視窗
	Sleep, 100 ; 稍微等一下
	Send, {Home}%ButtonText% ; 輸入文字
	return

Label_GoEdit:
	; 自動打開記事本並帶入檔名
	Run, notepad.exe "text_options.txt"
	Return

Label_GoAdd:
	; 建立一個新的小視窗
	Gui, Destroy
	Gui, 2:Default
	Gui, Add, Text, , 請輸入或貼上要新增的文字：
	Gui, Add, Edit, vNewText w250 h100
	Gui, Add, Button, gLabel_AddText w60 h30, 確定
	Gui, Add, Button, gLabel_CloseAdd w60 h30, 取消
	Gui, Show, , 新增文字
	return

Label_AddText:
	; 取得輸入的文字
	GuiControlGet, NewText

	; 寫入檔案並換行
	FileEncoding, UTF-8
	FileAppend, `n%NewText%, text_options.txt

	; 關閉新增視窗
	Gui, Destroy

	; 重新整理主選單
	GoSub, Label_GoTextGui
	Return

Label_CloseAdd:
	Gui, Destroy
	GoSub, Label_GoTextGui
	Return

; GUI關閉時的動作
Label_GuiClose:
	Gui, Destroy
	Return

;==================快捷鍵說明模組功能==================
Label_快捷鍵說明:
	; 檢查視窗是否已經存在，如果存在就顯示
	Gui, Destroy

	; 建立一個轉換函式，將符號轉成文字
	TransformHotkeySymbol(key) {
		hotkey_str := key
		display_str := ""
		
		; 檢查並處理修飾鍵
		if InStr(hotkey_str, "^")
		{
			display_str .= "CTRL + "
			hotkey_str := StrReplace(hotkey_str, "^")
		}
		
		if InStr(hotkey_str, "+")
		{
			display_str .= "SHIFT + "
			hotkey_str := StrReplace(hotkey_str, "+")
		}
		
		if InStr(hotkey_str, "!")
		{
			display_str .= "ALT + "
			hotkey_str := StrReplace(hotkey_str, "!")
		}
		
		if InStr(hotkey_str, "#")
		{
			display_str .= "WIN + "
			hotkey_str := StrReplace(hotkey_str, "#")
		}
		
		; 將剩下的主要按鍵轉成大寫
		StringUpper, hotkey_str, hotkey_str
		
		; 將剩下的主要按鍵加到最後
		display_str .= hotkey_str
		
		return display_str
	}

	; 取得轉換後的熱鍵文字
	Hotkey_結帳_顯示 := TransformHotkeySymbol(Hotkey_結帳)
	Hotkey_帶入客訂單_顯示 := TransformHotkeySymbol(Hotkey_帶入客訂單)
	Hotkey_直接列印發票_顯示 := TransformHotkeySymbol(Hotkey_直接列印發票)
	Hotkey_複製銷單_顯示 := TransformHotkeySymbol(Hotkey_複製銷單)
	Hotkey_緊急停止_顯示 := TransformHotkeySymbol(Hotkey_緊急停止)
	Hotkey_快速輸入_顯示 := TransformHotkeySymbol(Hotkey_快速輸入)
	Hotkey_快捷鍵說明_顯示 := TransformHotkeySymbol(Hotkey_快捷鍵說明)
	Hotkey_修改熱鍵_顯示 := TransformHotkeySymbol(Hotkey_修改熱鍵)
	Hotkey_全域設定_顯示 := TransformHotkeySymbol(Hotkey_全域設定) ; 取得全域設定熱鍵文字

	; 建立視窗
	Gui, Add, Text, x10 y10, 結帳
	Gui, Add, Text, x130 y10, %Hotkey_結帳_顯示%
	Gui, Add, Text, x10 y30, 帶入客訂單
	Gui, Add, Text, x130 y30, %Hotkey_帶入客訂單_顯示%
	Gui, Add, Text, x10 y50, 直接列印發票
	Gui, Add, Text, x130 y50, %Hotkey_直接列印發票_顯示%
	Gui, Add, Text, x10 y70, 複製銷單
	Gui, Add, Text, x130 y70, %Hotkey_複製銷單_顯示%
	Gui, Add, Text, x10 y90, 緊急停止
	Gui, Add, Text, x130 y90, %Hotkey_緊急停止_顯示%
	Gui, Add, Text, x10 y110, 快速輸入
	Gui, Add, Text, x130 y110, %Hotkey_快速輸入_顯示%
	Gui, Add, Text, x10 y130, 快捷鍵說明
	Gui, Add, Text, x130 y130, %Hotkey_快捷鍵說明_顯示%
	Gui, Add, Text, x10 y150, 修改熱鍵
	Gui, Add, Text, x130 y150, %Hotkey_修改熱鍵_顯示%
	Gui, Add, Text, x10 y170, 全域設定
	Gui, Add, Text, x130 y170, %Hotkey_全域設定_顯示%
	Gui, Add, Button, x10 y190 w200 h30 gLabel_GuiClose, 關閉
	Gui, Show, w250, 結帳小工具
	
	Return
	
;==================緊急停止功能模組==================
Label_緊急停止:
	; 這個熱鍵是緊急停止，不受開關限制，隨時都能執行。
	; 立即關閉所有執行中的熱鍵功能並重新啟動程式。
	MsgBox, 4624, 結帳小工具, 手動取消，重新載入小工具
	is_running_flag := 0
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
    Return

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
    Return

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
        ControlClick, Edit38, ahk_class ThunderRT6MDIForm,,,, NA
        ControlSetText, Edit38,, ahk_class ThunderRT6MDIForm
        ControlSendRaw, Edit38, 308`n, ahk_class ThunderRT6MDIForm
    }
    else if InStr(OutputVar1, Needlep2) {
        ControlClick, Edit38, ahk_class ThunderRT6MDIForm,,,, NA
        ControlSetText, Edit38,, ahk_class ThunderRT6MDIForm
        ControlSendRaw, Edit38, 312`n, ahk_class ThunderRT6MDIForm
    }
    else if InStr(OutputVar1, Needlek1) {
        ControlClick, Edit38, ahk_class ThunderRT6MDIForm,,,, NA
        ControlSetText, Edit38,, ahk_class ThunderRT6MDIForm
        ControlSendRaw, Edit38, 352`n, ahk_class ThunderRT6MDIForm
    }
    else if InStr(OutputVar1, Needlem1) {
        ControlClick, Edit38, ahk_class ThunderRT6MDIForm,,,, NA
        ControlSetText, Edit38,, ahk_class ThunderRT6MDIForm
        ControlSendRaw, Edit38, 394`n, ahk_class ThunderRT6MDIForm
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
                ControlGet, OutputVar, Visible,, Edit91, ahk_class ThunderRT6MDIForm
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
    MsgBox, 262704, 結帳小工具, 手動取消，請重新使用熱鍵！
	; 關閉開關，並重新載入程式。
	is_running_flag := 0
    reload
    Return

;==================檢查銷貨單狀態功能==================
main1:
    Gosub, slip
    ToolTip, 判斷銷貨單為新增狀態中..., 900, 300
    loop {
        ToolTip, 等待銷單為新增狀態中...., 900, 300
        ControlGet, OutputVar, Visible,, Edit91, ahk_class ThunderRT6MDIForm
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
                        ControlGet, OutputVar, Visible,, Edit91, ahk_class ThunderRT6MDIForm
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
    ControlGet, OutputVar, Visible,, Edit91, ahk_class ThunderRT6MDIForm
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

    ToolTip, 發票流程完成, 900, 300

    Return
	
;================== 熱鍵修改功能模組 ==================
Label_修改熱鍵:
	Gui, Destroy
	
	; 建立熱鍵修改視窗
	Gui, Add, Text, , 請選擇要修改的熱鍵：
	Gui, Add, DropDownList, vHotkeyName gLabel_UpdateHotkey, 結帳|帶入客訂單|直接列印發票|複製銷單|緊急停止|快速輸入|快捷鍵說明|修改熱鍵|全域設定 ; 新增全域設定選項
	Gui, Add, Text, x10 y60, 目前熱鍵：
	Gui, Add, Edit, x100 y60 w120 vCurrentHotkey ReadOnly
	Gui, Add, Text, x10 y90, 輸入新熱鍵：
	Gui, Add, Hotkey, x100 y90 w120 vNewHotkey
	Gui, Add, Button, x10 y120 w80 h30 gLabel_SaveHotkey, 儲存
	Gui, Add, Button, x100 y120 w80 h30 gLabel_GuiClose, 取消
	Gui, Show, w240, 修改熱鍵
	
	; 預設顯示第一個選項的熱鍵
	GuiControl, Choose, HotkeyName, 1
	GoSub, Label_UpdateHotkey
	return

Label_UpdateHotkey:
	Gui, Submit, NoHide
	
	; 根據選擇的項目，取得對應的熱鍵變數
	if (HotkeyName = "結帳")
		hotkey_to_display := Hotkey_結帳
	else if (HotkeyName = "帶入客訂單")
		hotkey_to_display := Hotkey_帶入客訂單
	else if (HotkeyName = "直接列印發票")
		hotkey_to_display := Hotkey_直接列印發票
	else if (HotkeyName = "複製銷單")
		hotkey_to_display := Hotkey_複製銷單
	else if (HotkeyName = "緊急停止")
		hotkey_to_display := Hotkey_緊急停止
	else if (HotkeyName = "快速輸入")
		hotkey_to_display := Hotkey_快速輸入
	else if (HotkeyName = "快捷鍵說明")
		hotkey_to_display := Hotkey_快捷鍵說明
	else if (HotkeyName = "修改熱鍵")
		hotkey_to_display := Hotkey_修改熱鍵
	else if (HotkeyName = "全域設定")
		hotkey_to_display := Hotkey_全域設定 ; 讀取全域設定熱鍵

	; 呼叫函式將符號熱鍵轉成文字
	display_text := TransformHotkeySymbol(hotkey_to_display)
	
	; 更新 GUI 上的熱鍵文字
	GuiControl,, CurrentHotkey, %display_text%

	return

Label_SaveHotkey:
	Gui, Submit
	
	if (NewHotkey = "")
	{
		MsgBox, 0, 錯誤, 新熱鍵不能為空白！
		Return
	}

	; 解除舊熱鍵的綁定
	Hotkey, %Hotkey_結帳%, Off
	Hotkey, %Hotkey_帶入客訂單%, Off
	Hotkey, %Hotkey_直接列印發票%, Off
	Hotkey, %Hotkey_複製銷單%, Off
	Hotkey, %Hotkey_緊急停止%, Off
	Hotkey, %Hotkey_快速輸入%, Off
	Hotkey, %Hotkey_快捷鍵說明%, Off
	Hotkey, %Hotkey_修改熱鍵%, Off
	Hotkey, %Hotkey_全域設定%, Off ; 解除全域設定舊熱鍵的綁定

	; 根據選擇更新對應的熱鍵變數
	if (HotkeyName = "結帳")
		Hotkey_結帳 := NewHotkey
	else if (HotkeyName = "帶入客訂單")
		Hotkey_帶入客訂單 := NewHotkey
	else if (HotkeyName = "直接列印發票")
		Hotkey_直接列印發票 := NewHotkey
	else if (HotkeyName = "複製銷單")
		Hotkey_複製銷單 := NewHotkey
	else if (HotkeyName = "緊急停止")
		Hotkey_緊急停止 := NewHotkey
	else if (HotkeyName = "快速輸入")
		Hotkey_快速輸入 := NewHotkey
	else if (HotkeyName = "快捷鍵說明")
		Hotkey_快捷鍵說明 := NewHotkey
	else if (HotkeyName = "修改熱鍵")
		Hotkey_修改熱鍵 := NewHotkey
	else if (HotkeyName = "全域設定")
		Hotkey_全域設定 := NewHotkey ; 更新全域設定熱鍵變數

    ; 將新的熱鍵設定寫入檔案
    IniWrite, %Hotkey_結帳%, Hotkeys.ini, Hotkeys, 結帳
    IniWrite, %Hotkey_帶入客訂單%, Hotkeys.ini, Hotkeys, 帶入客訂單
    IniWrite, %Hotkey_直接列印發票%, Hotkeys.ini, Hotkeys, 直接列印發票
    IniWrite, %Hotkey_複製銷單%, Hotkeys.ini, Hotkeys, 複製銷單
    IniWrite, %Hotkey_緊急停止%, Hotkeys.ini, Hotkeys, 緊急停止
    IniWrite, %Hotkey_快速輸入%, Hotkeys.ini, Hotkeys, 快速輸入
    IniWrite, %Hotkey_快捷鍵說明%, Hotkeys.ini, Hotkeys, 快捷鍵說明
    IniWrite, %Hotkey_修改熱鍵%, Hotkeys.ini, Hotkeys, 修改熱鍵
    IniWrite, %Hotkey_全域設定%, Hotkeys.ini, Hotkeys, 全域設定 ; 儲存全域設定熱鍵
    
    ; 在熱鍵儲存後，將 RunStatus 的值設為 0
    IniWrite, 0, Hotkeys.ini, Settings, RunStatus

    ; 重新綁定所有熱鍵
    Hotkey, %Hotkey_結帳%, Label_結帳
    Hotkey, %Hotkey_帶入客訂單%, Label_帶入客訂單
    Hotkey, %Hotkey_直接列印發票%, Label_直接列印發票
    Hotkey, %Hotkey_複製銷單%, Label_複製銷單
    Hotkey, %Hotkey_緊急停止%, Label_緊急停止
    Hotkey, %Hotkey_快速輸入%, Label_快速輸入
    Hotkey, %Hotkey_快捷鍵說明%, Label_快捷鍵說明
    Hotkey, %Hotkey_修改熱鍵%, Label_修改熱鍵
    Hotkey, %Hotkey_全域設定%, Label_全域設定 ; 重新綁定全域設定熱鍵
    
    ; 關閉視窗並提示
    Gui, Destroy
    MsgBox, 0, 成功, 熱鍵已修改成功！下次啟動時將會再次顯示熱鍵說明。
	Reload
	Return
	
;==================全域設定模組==================
Label_全域設定:
	; 檢查開關，如果正在執行中，就立即停止。
	if is_running_flag {
		Return
	}
	; 打開開關，表示功能正在執行。
	is_running_flag := 1
	
	; 檢查視窗是否已經存在，如果存在就顯示
	Gui, Destroy

	; 建立全域設定視窗
	Gui, Add, Text, , 請設定 OSD 顯示：
	; 根據 OSD_enabled 的設定來決定勾選狀態
	if (OSD_enabled = 1) {
		Gui, Add, Checkbox, vOSD_Setting Checked, 開啟 OSD
	} else {
		Gui, Add, Checkbox, vOSD_Setting, 開啟 OSD
	}
	
	Gui, Add, Button, gLabel_SaveSettings w80 h30, 儲存
	Gui, Add, Button, gLabel_GuiClose w80 h30, 取消
	Gui, Show, w200, 全域設定
	Return

Label_SaveSettings:
	Gui, Submit, NoHide
	
	; 讀取 checkbox 的狀態
	if OSD_Setting {
		OSD_enabled := 1
	} else {
		OSD_enabled := 0
	}
	
	; 將新設定寫入檔案
	IniWrite, %OSD_enabled%, Hotkeys.ini, Settings, OSD

	; 關閉視窗並提示
	Gui, Destroy
	MsgBox, 0, 成功, 全域設定已儲存！
	
	; 關閉開關
	is_running_flag := 0
	Return