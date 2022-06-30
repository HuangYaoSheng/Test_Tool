#RequireAdmin
#include <ScreenCapture.au3>
#include <File.au3>

;
; AutoHDBench_A0 Design by YS.Huang
;
Local $ToolTitle = "HDBENCH"
Local $ToolPath = "./hdb340b6\"

FileDelete("./Result_Log.log")
FileDelete("./Result_Pic.jpg")

$OpenTool = ShellExecute($ToolTitle, "", $ToolPath)

If $OpenTool = 0 Then	
	MsgBox(266144, "Error", "Open HDBENCH Tool Error")	
Else
	
	WinWait($ToolTitle, "", 3)
	WinActivate($ToolTitle, "")
	WinWaitActive($ToolTitle, "", 3)

	$RunDisk = IniRead("./Setup.ini", "Setting", "Disk", "")
	$RunPattern = IniRead("./Setup.ini", "Setting", "Pattern", "")

	;
	; Select test disk and test patten
	;
	ControlCommand($ToolTitle,"","ComboBox1","SelectString",$RunDisk)
	Sleep(500)
	ControlCommand($ToolTitle,"","ComboBox2","SelectString",$RunPattern)
	Sleep(500)

	WinWait($ToolTitle, "", 3)
	WinActivate($ToolTitle, "")
	WinWaitActive($ToolTitle, "", 3)

	;
	; Start test
	;
	ControlClick($ToolTitle,"","Button16")
	Sleep(500)

	;
	; Wait test status
	;
	$state = WinGetState($ToolTitle, "")
	$CheckIsEnabled = BitAnd($state, 2) ; 視窗可見 = 2 / 不可見 = 0

	While $CheckIsEnabled == 0	
		$state = WinGetState($ToolTitle, "")
		$CheckIsEnabled = BitAnd($state, 2)
		Sleep(5000)	
	WEnd
		Sleep(3000)
		
	;
	; Save Test Log to ./
	;
	WinWait($ToolTitle, "", 3)
	WinActivate($ToolTitle, "")
	WinWaitActive($ToolTitle, "", 3)

	ControlClick($ToolTitle, "", "[CLASS:Button; INSTANCE:24]", "Left", 3, 41, 15) ; 滑鼠左鍵點擊copy按鈕
	Sleep(500)

	$File = FileOpen("./Result_Log.log", 2)
	$Log = ClipGet()
	Sleep(500)	
	_FileWriteLog(@ScriptDir & "\Result_Log.log", $Log)
	Sleep(500)
	
	FileClose($File)
	
	;
	; Save Test Tool Pic to ./
	;
	Sleep(500)
	_ScreenCapture_CaptureWnd(@ScriptDir & "\Result_Pic.jpg", $ToolTitle) ; 儲存測試視窗
	
	Sleep(500)
	WinClose($ToolTitle)
	
EndIf







