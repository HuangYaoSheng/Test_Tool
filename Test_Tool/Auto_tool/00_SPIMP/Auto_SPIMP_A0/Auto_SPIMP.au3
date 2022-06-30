#RequireAdmin
#include <file.au3>
#include <ScreenCapture.au3> 

;
; Auto_SPIMP_A0 Design by YS.Huang
;
Local $ToolTitle = "MyTek MP Tool SPI"
Local $ToolExe = "MyTek SPIMP.exe"
Local $ToolPath = "./dll\"
Local $CheckIsScan = 0
Local $CheckIsStart = 0
Local $CompareToolTXT = "MyTek SPI MP Tool - Port information"
Local $PassStr_1 = 0, $PassStr_2 = 0, $PassStr_3 = 0, $PassStr_4 = 0, $PassStr_5 = 0, $PassStr_6 = 0, $PassStr_7 = 0, $PassStr_8 = 0, $PassStr_9 = 0, $PassStr_10 = 0, $PassStr_11 = 0
Local  $SetupPath = "./Setup.ini"
Local $ReadIni = ""
Local $PassStr = "$PassStr_"
Local $CompareIniResult = ""
Local $SelectPort = "Null"
Local $FlowSetting = ""

;
; Main test flow
;
DirRemove("./Test_Log", 1)
DirCreate("./Test_Log")
FileDelete("./Test_Result_Log.csv")
$OpenTool = ShellExecute($ToolExe, "", $ToolPath)
	
	If $OpenTool = 0 Then	
		MsgBox(266144, "Error", "Open MyTek SPIMP.exe Error")	
	Else
		WinWait($ToolTitle, "", 3)
		WinActivate($ToolTitle, "")
		WinWaitActive($ToolTitle, "", 3)
		WinMove($ToolTitle, "", 0, 0)
	EndIf
	
$ReadIniTestTimes = IniRead($SetupPath, "InitialSetting", "TestTimes", "")
$ReadIniFlow = IniRead($SetupPath, "InitialSetting", "TestFlow", "")
$FileCSV = FileOpen("./Test_Result_Log.csv", 2)
FileWrite($FileCSV, " ,")
OptHead()
For $TestCut = 1 to $ReadIniTestTimes Step +1
	If $TestCut <> 1 Then
		FileWrite($FileCSV, @CRLF & $TestCut & ",")
	Else
		FileWrite($FileCSV, $TestCut & ",")
	EndIf
	
	If $ReadIniFlow = 1 Then ScanFlow()
	If $ReadIniFlow = 2 Then StartFlow()
	If $ReadIniFlow = 3 Then 
		ScanFlow()
		FileWrite($FileCSV, @CRLF & ",")
		StartFlow()
	EndIf
Next
MsgBox(266144, "Auto_SPIMP", "Test Done")

;
; Function items
;
Func OptHead()
	Local $Check_1_8 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:3]", "IsChecked", "")
	Local $Check_9_16 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:4]", "IsChecked", "")
	Local $Check_17_24 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:5]", "IsChecked", "")
	Local $Check_25_32 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:6]", "IsChecked", "")
	Local $Check_33_40 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:15]", "IsChecked", "")	
	If $Check_1_8 = 1 Then 
		$OptHead_1_8 = "Port1,Port2,Port3,Port4,Port5,Port6,Port7,Port8,"
		FileWriteLine($FileCSV, $OptHead_1_8)
	EndIf
	If $Check_9_16 = 1 Then 
		$OptHead_9_16 = "Port9,Port10,Port11,Port12,Port13,Port14,Port15,Port16,"
		FileWriteLine($FileCSV, $OptHead_9_16)
	EndIf
	If $Check_17_24 = 1 Then
		$OptHead_17_24 = "Port17,Port18,Port19,Port20,Port21,Port22,Port23,Port24,"
		FileWriteLine($FileCSV, $OptHead_17_24)
	EndIf
	If $Check_25_32 = 1 Then
		$OptHead_25_32 = "Port25,Port26,Port27,Port28,Port29,Port30,Port31,Port32,"
		FileWriteLine($FileCSV, $OptHead_25_32)
	EndIf
	If $Check_33_40 = 1 Then
		$OptHead_23_40 = "Port33,Port34,Port35,Port36,Port37,Port38,Port39,Port40,"
		FileWriteLine($FileCSV, $OptHead_23_40)
	EndIf
EndFunc

Func ScanFlow()
	$FlowSetting = 1
	WinWait($ToolTitle, "", 3)
	WinActivate($ToolTitle, "")
	WinWaitActive($ToolTitle, "", 3)
	ControlClick($ToolTitle, "", "[CLASS:SysTabControl32; INSTANCE:1]", "left", 1, 26, 12) ; Click SPI Sheet
	Sleep(500)
	Auto_Click_Scan()	
	Local $Check_1_8 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:3]", "IsChecked", "")
	Local $Check_9_16 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:4]", "IsChecked", "")
	Local $Check_17_24 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:5]", "IsChecked", "")
	Local $Check_25_32 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:6]", "IsChecked", "")
	Local $Check_33_40 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:15]", "IsChecked", "")
	If $Check_1_8 = 1 Then Check_Result_1_8()
	If $Check_9_16 = 1 Then Check_Result_9_16()
	If $Check_17_24 = 1 Then Check_Result_17_24()
	If $Check_25_32 = 1 Then Check_Result_25_32 ()
	If $Check_33_40 = 1 Then Check_Result_33_40 ()
EndFunc

Func StartFlow()
	$FlowSetting = 2
	WinWait($ToolTitle, "", 3)
	WinActivate($ToolTitle, "")
	WinWaitActive($ToolTitle, "", 3)	
	ControlClick($ToolTitle, "", "[CLASS:SysTabControl32; INSTANCE:1]", "left", 1, 26, 12) ; Click SPI Sheet
	Sleep(500)		
	Auto_Click_Start()
	Local $Check_1_8 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:3]", "IsChecked", "")
	Local $Check_9_16 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:4]", "IsChecked", "")
	Local $Check_17_24 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:5]", "IsChecked", "")
	Local $Check_25_32 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:6]", "IsChecked", "")
	Local $Check_33_40 = ControlCommand($ToolTitle, "", "[CLASS:Button; INSTANCE:15]", "IsChecked", "")
	If $Check_1_8 = 1 Then Check_Result_1_8()
	If $Check_9_16 = 1 Then Check_Result_9_16()
	If $Check_17_24 = 1 Then Check_Result_17_24()
	If $Check_25_32 = 1 Then Check_Result_25_32 ()
	If $Check_33_40 = 1 Then Check_Result_33_40 ()
EndFunc

Func Auto_Click_Scan()
	Local $CheckIsScan = 0
	WinWait($ToolTitle, "", 3)
	WinActivate($ToolTitle, "")
	WinWaitActive($ToolTitle, "", 3)
	ControlClick($ToolTitle, "Scan", "[CLASS:Button; INSTANCE:7]")
	Sleep(2000)
	While $CheckIsScan == 0
		$CheckIsScan = ControlCommand($ToolTitle, "Scan", "[CLASS:Button; INSTANCE:7]", "IsEnabled", "") ; Scan Button Enable = 1 ; Disable = 0
		Sleep(1000)
	WEnd
		Sleep(3000)
EndFunc

Func Auto_Click_Start()
	Local $CheckIsStart = 0
	WinWait($ToolTitle, "", 3)
	WinActivate($ToolTitle, "")
	WinWaitActive($ToolTitle, "", 3)
	ControlClick($ToolTitle, "Start", "[CLASS:Button; INSTANCE:1]")
	Sleep(2000)
	While $CheckIsStart == 0
		$CheckIsStart = ControlCommand($ToolTitle, "Start", "[CLASS:Button; INSTANCE:1]", "IsEnabled", "") ; Start Button Enable = 1 ; Disable = 0
		Sleep(1000)
	WEnd
		Sleep(3000)
EndFunc
	
Func Compare_Scan()
	WinWait($CompareToolTXT, "", 3)
	WinActivate($CompareToolTXT, "")
	WinWaitActive($CompareToolTXT, "", 3)
	For $CompareScanIniCut = 1 to 11 Step +1
		$ReadIni = IniRead($SetupPath, "ScanPassWard", "PassStr_" & $CompareScanIniCut, "")
		$CompareIniResult = StringInStr(WinGetText($CompareToolTXT), $ReadIni, 1) ; Fail = 0 ; Pass > 0
	Next
	If $CompareIniResult = 0 Then
		$FileCSV = FileOpen("./Test_Result_Log.csv", 1)	
		FileWrite($FileCSV, "Scan Fail,")
		$LogData = Fileopen("./Test_Log\" & $SelectPort & "_Scan" & $TestCut & "_Fail" & ".Log", 1)
		FileWrite($LogData, ControlGetText($CompareToolTXT, "", "Edit1"))
	Else
		$FileCSV = FileOpen("./Test_Result_Log.csv", 1)
		FileWrite($FileCSV, "Scan Pass,")
		$LogData = Fileopen("./Test_Log\" & $SelectPort & "_Scan" & $TestCut & "_Pass" & ".Log", 1)
		FileWrite($LogData, ControlGetText($CompareToolTXT, "", "Edit1"))
	EndIf
	$handle = FileOpen("./Test_Result_Log.csv", 1)
	FileClose($handle)
	WinClose($CompareToolTXT)
EndFunc

Func Compare_Start()
	WinWait($CompareToolTXT, "", 3)
	WinActivate($CompareToolTXT, "")
	WinWaitActive($CompareToolTXT, "", 3)
	For $CompareStartIniCut = 1 to 11 Step +1
		$ReadIni = IniRead($SetupPath, "StartPassWard", "PassStr_" & $CompareStartIniCut, "")
		$CompareIniResult = StringInStr(WinGetText($CompareToolTXT), $ReadIni, 1) ; Fail = 0 ; Pass > 0
	Next
	If $CompareIniResult = 0 Then
		$FileCSV = FileOpen("./Test_Result_Log.csv", 1)	
		FileWrite($FileCSV, "Initial Fail,")
		$LogData = Fileopen("./Test_Log\" & $SelectPort & "_Initial" & $TestCut & "_Fail" & ".Log", 1)
		FileWrite($LogData, ControlGetText($CompareToolTXT, "", "Edit1"))
	Else
		$FileCSV = FileOpen("./Test_Result_Log.csv", 1)
		FileWrite($FileCSV, "Initial Pass,")
		$LogData = Fileopen("./Test_Log\" & $SelectPort & "_Initial" & $TestCut & "_Pass" & ".Log", 1)
		FileWrite($LogData, ControlGetText($CompareToolTXT, "", "Edit1"))
	EndIf
	$handle = FileOpen("./Test_Result_Log.csv", 1)
	FileClose($handle)
	WinClose($CompareToolTXT)
EndFunc

Func PortXY_1()
	MouseClick("left", 81, 219)
	$SelectPort ="Port1"
EndFunc

Func PortXY_2()
	MouseClick("left", 193, 221)
	$SelectPort ="Port2"
EndFunc

Func PortXY_3()
	MouseClick("left", 302, 222)
	$SelectPort ="Port3"
EndFunc

Func PortXY_4()
	MouseClick("left", 411, 220)
	$SelectPort ="Port4"
EndFunc

Func PortXY_5()
	MouseClick("left", 520, 222)
	$SelectPort ="Port5"
EndFunc

Func PortXY_6()
	MouseClick("left", 631, 221)
	$SelectPort ="Port6"
EndFunc

Func PortXY_7()
	MouseClick("left", 740, 222)
	$SelectPort ="Port7"
EndFunc

Func PortXY_8()
	MouseClick("left", 851, 222)
	$SelectPort ="Port8"
EndFunc

Func PortXY_9()
	MouseClick("left", 81, 324)
	$SelectPort ="Port9"
EndFunc

Func PortXY_10()
	MouseClick("left", 192, 324)
	$SelectPort ="Port10"
EndFunc

Func PortXY_11()
	MouseClick("left", 303, 326)
	$SelectPort ="Port11"
EndFunc

Func PortXY_12()
	MouseClick("left", 411, 324)
	$SelectPort ="Port12"
EndFunc

Func PortXY_13()
	MouseClick("left", 523, 325)
	$SelectPort ="Port13"
EndFunc

Func PortXY_14()
	MouseClick("left", 631, 324)
	$SelectPort ="Port14"
EndFunc

Func PortXY_15()
	MouseClick("left", 742, 327)
	$SelectPort ="Port15"
EndFunc

Func PortXY_16()
	MouseClick("left", 851, 324)
	$SelectPort ="Port16"
EndFunc

Func PortXY_17()
	MouseClick("left", 82, 432)
	$SelectPort ="Port17"
EndFunc

Func PortXY_18()
	MouseClick("left", 191, 432)
	$SelectPort ="Port18"
EndFunc

Func PortXY_19()
	MouseClick("left", 302, 433)
	$SelectPort ="Port19"
EndFunc

Func PortXY_20()
	MouseClick("left", 410, 434)
	$SelectPort ="Port20"
EndFunc

Func PortXY_21()
	MouseClick("left", 522, 431)
	$SelectPort ="Port21"
EndFunc

Func PortXY_22()
	MouseClick("left", 631, 433)
	$SelectPort ="Port22"
EndFunc

Func PortXY_23()
	MouseClick("left", 741, 431)
	$SelectPort ="Port23"
EndFunc

Func PortXY_24()
	MouseClick("left", 852, 433)
	$SelectPort ="Port24"
EndFunc

Func PortXY_25()
	MouseClick("left", 82, 536)
	$SelectPort ="Port25"
EndFunc

Func PortXY_26()
	MouseClick("left", 191, 537)
	$SelectPort ="Port26"
EndFunc

Func PortXY_27()
	MouseClick("left", 301, 537)
	$SelectPort ="Port27"
EndFunc

Func PortXY_28()
	MouseClick("left", 410, 538)
	$SelectPort ="Port28"
EndFunc

Func PortXY_29()
	MouseClick("left", 521, 537)
	$SelectPort ="Port29"
EndFunc

Func PortXY_30()
	MouseClick("left", 632, 537)
	$SelectPort ="Port30"
EndFunc

Func PortXY_31()
	MouseClick("left", 742, 538)
	$SelectPort ="Port31"
EndFunc

Func PortXY_32()
	MouseClick("left", 851, 536)
	$SelectPort ="Port32"
EndFunc

Func PortXY_33()
	MouseClick("left", 83, 644)
	$SelectPort ="Port33"
EndFunc

Func PortXY_34()
	MouseClick("left", 191, 643)
	$SelectPort ="Port34"
EndFunc

Func PortXY_35()
	MouseClick("left", 302, 642)
	$SelectPort ="Port35"
EndFunc

Func PortXY_36()
	MouseClick("left", 410, 644)
	$SelectPort ="Port36"
EndFunc

Func PortXY_37()
	MouseClick("left", 521, 642)
	$SelectPort ="Port37"
EndFunc

Func PortXY_38()
	MouseClick("left", 632, 642)
	$SelectPort ="Port38"
EndFunc

Func PortXY_39()
	MouseClick("left", 742, 641)
	$SelectPort ="Port39"
EndFunc

Func PortXY_40()
	MouseClick("left", 851, 643)
	$SelectPort ="Port40"
EndFunc

Func Check_Result_1_8()
	BlockInput(1)
	ScreenCapPic()
	PortXY_1()
		CompareMode()
	PortXY_2()
		CompareMode()
	PortXY_3()
		CompareMode()
	PortXY_4()
		CompareMode()
	PortXY_5()
		CompareMode()
	PortXY_6()
		CompareMode()
	PortXY_7()
		CompareMode()
	PortXY_8()
		CompareMode()
	BlockInput(0)
EndFunc

Func Check_Result_9_16()
	BlockInput(1)
	ScreenCapPic()
	PortXY_9()
		CompareMode()
	PortXY_10()
		CompareMode()
	PortXY_11()
		CompareMode()
	PortXY_12()
		CompareMode()
	PortXY_13()
		CompareMode()
	PortXY_14()
		CompareMode()
	PortXY_15()
		CompareMode()
	PortXY_16()
		CompareMode()
	BlockInput(0)
EndFunc

Func Check_Result_17_24()
	BlockInput(1)
	ScreenCapPic()
	PortXY_17()
		CompareMode()
	PortXY_18()
		CompareMode()
	PortXY_19()
		CompareMode()
	PortXY_20()
		CompareMode()
	PortXY_21()
		CompareMode()
	PortXY_22()
		CompareMode()
	PortXY_23()
		CompareMode()
	PortXY_24()
		CompareMode()
	BlockInput(0)
EndFunc

Func Check_Result_25_32()
	BlockInput(1)
	ScreenCapPic()
	PortXY_25()
		CompareMode()
	PortXY_26()
		CompareMode()
	PortXY_27()
		CompareMode()
	PortXY_28()
		CompareMode()
	PortXY_29()
		CompareMode()
	PortXY_30()
		CompareMode()
	PortXY_31()
		CompareMode()
	PortXY_32()
		CompareMode()
	BlockInput(0)
EndFunc

Func Check_Result_33_40()
	BlockInput(1)
	ScreenCapPic()
	PortXY_33()
		CompareMode()
	PortXY_34()
		CompareMode()
	PortXY_35()
		CompareMode()
	PortXY_36()
		CompareMode()
	PortXY_37()
		CompareMode()
	PortXY_38()
		CompareMode()
	PortXY_39()
		CompareMode()
	PortXY_40()
		CompareMode()
	BlockInput(0)
EndFunc
	
Func CompareMode()
	If $FlowSetting = 1 Then 
		Compare_Scan()
	EndIf
	If $FlowSetting = 2 Then
		Compare_Start()
	EndIf
EndFunc

Func ScreenCapPic()
	WinWait($ToolTitle, "", 3)
	WinActivate($ToolTitle, "")
	WinWaitActive($ToolTitle, "", 3)
	If $FlowSetting = 1 Then _ScreenCapture_CaptureWnd(@ScriptDir & "./Test_Log\Scan" & $TestCut & "_Result_Pic.jpg", $ToolTitle)
	If $FlowSetting = 2 Then _ScreenCapture_CaptureWnd(@ScriptDir & "./Test_Log\Start" & $TestCut & "_Result_Pic.jpg", $ToolTitle)
EndFunc


















