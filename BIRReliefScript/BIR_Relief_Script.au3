#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=res\icon\ScriptIcon.ico
#AutoIt3Wrapper_Res_Description=BIR Relief System Script
#AutoIt3Wrapper_Res_Fileversion=1.2.1
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; #include <ImageSearch2015.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <Misc.au3>
#include <WinAPI.au3>
#include <BlockInputEx.au3>
; #RequireAdmin

HotKeySet("{END}", "exitz")

AutoItSetOption("SendKeyDelay", 110) ; 100
AutoItSetOption("WinTitleMatchMode", 2)
; AutoItSetOption("GUIOnEventMode", 1)

; CONSTANTS
; Script mode constants
Local Const $M_SALES = 0
Local Const $M_PURCHASE = 1

; purchase mode constants
Local Const $P_SERVICES = 0
Local Const $P_CAPITAL_GOODS = 1
Local Const $P_OTHER_GOODS = 2

; files
Local $excelFile = ""

; WinTitles
Local Const $BIRtitle = "Bureau of Internal Revenue - Relief System"
Local $ExcelTitle = ""

; script states
Local $scriptMode = $M_SALES
Local $purchaseMode = $P_SERVICES
Local $isContinue = True
Local $isTest = False

; caret positions
; Local Const $CARET_POS_SALES =

; BIR Dialogs
; BIR System Message : [CLASS:#32770] - Name err
; Microsoft Visual FoxPro: [CLASS:#32770] - Exceed

; SCRIPT START

; check if another instance of script is already running
If _Singleton("BIR_Relief_Script", 1) = 0 Then
	MsgBox(16, "Error", "The Script is already running.", 3)
	Exit 0
EndIf

; MAINGUI
$frmMain = GUICreate("BIR System Script", 432, 280, 541, 168)
$mnuHelp = GUICtrlCreateMenu("&Help")
$itmAbout = GUICtrlCreateMenuItem("&About", $mnuHelp)
GUISetBkColor(0xFFFFFF)
$lblExcel = GUICtrlCreateLabel("Excel File Name:", 21, 16, 83, 17)
$txtFileName = GUICtrlCreateInput("", 112, 12, 273, 21)
$grp = GUICtrlCreateGroup("Purchase Mode", 16, 120, 401, 57)
$radServices = GUICtrlCreateRadio("Services", 40, 144, 73, 17)
$radCapitalGoods = GUICtrlCreateRadio("Capital Goods", 158, 144, 97, 17)
$radOtherGoods = GUICtrlCreateRadio("Other Goods", 322, 144, 89, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$btnStart = GUICtrlCreateButton("Start!", 256, 209, 75, 25)
$btnCancel = GUICtrlCreateButton("Cancel", 344, 209, 75, 25)
$chkTest = GUICtrlCreateCheckbox("Test", 40, 216, 49, 17)
$radPurchase = GUICtrlCreateRadio("Purchase", 16, 96, 73, 17)
$radSales = GUICtrlCreateRadio("Sales", 16, 72, 57, 17)
$Label1 = GUICtrlCreateLabel("Choose what to automate:", 16, 48, 128, 17)
$btnBrowse = GUICtrlCreateButton("...", 390, 12, 27, 21)
$Label2 = GUICtrlCreateLabel("Options:", 16, 192, 43, 17)

; default states
GUICtrlSetState($radServices, $GUI_CHECKED)
GUICtrlSetState($radSales, $GUI_CHECKED)
enableRadioServices(False)

; tooltips
GUICtrlSetTip($chkTest, "If checked, the script will not save the records to the BIR System.")
GUICtrlSetTip($btnBrowse, "Browse the Excel file from your PC.")

GUISetState(@SW_SHOW)


While 1
	Switch GUIGetMsg()
		Case $btnBrowse
			browse()

		Case $btnStart
			start()

		Case $radSales
			enableRadioServices(False)
			$scriptMode = $M_SALES

		Case $radPurchase
			enableRadioServices(True)
			$scriptMode = $M_PURCHASE

		Case $itmAbout
			itmAbout()

		Case $chkTest
			If GUICtrlRead($chkTest) = $GUI_CHECKED Then
				$isTest = True
			Else
				$isTest = False
			EndIf

		Case $GUI_EVENT_CLOSE, $btnCancel
			mainGUIClose()
	EndSwitch
WEnd
; END MAINGUI

; Description: will automate the purchase
Func automate()
	; go to excel
	activateWin($ExcelTitle, "Excel")

	While $isContinue
		MouseMove(0, 0, 5)

		Local $caretPos[2]
		Local $isTINexists = False

		; copy TIN
		Send("{HOME}^c")
		; check TIN if not empty, if empty exit
		Local $TIN = ClipGet()
		$TIN = StringStripCR($TIN)
		$TIN = StringStripWS($TIN, 8)
		; check if reached the end
		If (StringLen($TIN) = 0) Or ($TIN = "GrandTotal") Then
			_BlockInputEx(0)
			MsgBox(64, "Finished", "Script has reached the end of the records.")
			exitz()
		EndIf
		;MsgBox(0, "", "\'" & $TIN & "\'")

		; go to BIR
		activateWin($BIRtitle)
		; remove revert
		Send("!r")
		; add new
		Send("!a")
		Send("{DOWN}{UP}^v{TAB}")


		Sleep(30)
		; get caret position to check if TIN already exists
		Local $caretPos = WinGetCaretPos()
		; MsgBox(0, "", $caretPos[0] & ", " & $caretPos[1])

		if $caretPos[0] <> 16 Then
			$isTINexists = True
		EndIf

		; copy reg name and address
		if $isTINexists = False Then
			Local $copied = ""
			; copy reg name
			; 50 chr limit, '.
			activateWin($ExcelTitle, "Excel")
			Send("{HOME}{RIGHT}^c")
			$copied = ClipGet()
			fixString($copied)
			chkStringLimit($copied, 50)
			activateWin($BIRtitle)
			ClipPut($copied)
			Send("^v{TAB 4}")

			; copy add 1
			; 30 chr limit
			activateWin($ExcelTitle, "Excel")
			Send("{RIGHT}^c")
			$copied = ClipGet()
			fixString($copied)
			; if "(blank)" then paste manila
			if StringInStr($copied, "blank") <> 0 Then
				$copied = "MANILA"
			EndIf
			activateWin($BIRtitle)
			ClipPut($copied)
			Send("^v{TAB}")

			; copy add 2
			; 30 chr limit
			activateWin($ExcelTitle, "Excel")
			Send("{RIGHT}^c")
			$copied = ClipGet()
			fixString($copied)
			; if "(blank)" then paste manila
			if StringInStr($copied, "blank") <> 0 Then
				$copied = "MANILA"
			EndIf
			activateWin($BIRtitle)
			ClipPut($copied)
			Send("^v{TAB}")
		EndIf

		; copy taxable and services
		activateWin($ExcelTitle, "Excel")
		Send("^{RIGHT}{LEFT 2}^c")

		activateWin($BIRtitle)
		If $scriptMode = $M_PURCHASE Then
			Send("{TAB 2}")
		EndIf
		Send("^v{TAB}")

		; services
		Switch $scriptMode
			Case $M_SALES
				Send("{TAB 3}")

			Case $M_PURCHASE
				Switch $purchaseMode
					Case $P_SERVICES
						Send("{TAB 3}")
					case $P_CAPITAL_GOODS
						Send("{DELETE}{TAB 3}")
					case $P_OTHER_GOODS
						Send("{DELETE}{TAB}{DELETE}{TAB 2}")
				EndSwitch
		EndSwitch

		Send("12")

		; check if valid
		Sleep(20)
		Send("^c")
		Local $BIRtotal = ClipGet()
		activateWin($ExcelTitle, "Excel")
		Sleep(20)
		Send("^{RIGHT}^c")
		Local $ExcelTotal = ClipGet()

		; manipulate string
		; remove spaces and comma
		fixStringNumber($BIRtotal)
		fixStringNumber($ExcelTotal)

		if $BIRtotal <> $ExcelTotal Then
			_BlockInputEx(0)
			$msg = MsgBox(4, "Not Equal", "Total input Tax is not equal to Sum of TAX" & @CRLF & "BIR Total:        " & $BIRtotal & _
			@CRLF & "In Excel Total: " & $ExcelTotal & @CRLF & @CRLF & "Do you still want to  proceed?")
			if $msg = 7 Then
				exitz()
			EndIf
		EndIf

		_BlockInputEx(1, "{END}")
		; Save
		If Not $isTest Then
			activateWin($BIRtitle)
			Send("!s")
			activateWin($ExcelTitle, "Excel")

			; mark saved
			; Send("{RIGHT}D{ENTER}{HOME}")
			Send("{DOWN}{HOME}")
		Else
			Send("{DOWN}")
		EndIf

	WEnd
EndFunc ; automate


; SCRIPT END

; FUNCTIONS

; Description: searches the control of mode checks wether valid or not
Func getMode($M)
	Local $mode = ControlGetText("Bureau of Internal Revenue - Relief System", "", "[CLASS:birrlf16c000000; INSTANCE:4]")
	if @error = 1 Then
		Return False
	EndIf

	Switch $M
		Case $M_SALES
			If $mode = "Sales Data Entry Screen - REBTRADE INTERNATIONAL CORPORATION" Then
				Return True
			EndIf

		Case $M_PURCHASE
			If $mode = "Purchase Data Entry Screen - REBTRADE INTERNATIONAL CORPORATION" Then
				Return True
			EndIf
	EndSwitch
	Return False
EndFunc

; Description: check if the string exceeds the given limit
; Parameters:
; 		$s - String to check
; 		$l - the number of limit
Func chkStringLimit($s, $l)
	; count string
	Local $count = StringLen($s)
	If $count >= $l Then
		; show gui to change
	EndIf
EndFunc

; Description: remove ' . , and relplace & with 'and'
; Parameters:
; 		$s - the string to manipulate
Func fixString(ByRef $s)
	$s = StringReplace($s, ".", "")
	$s = StringReplace($s, ",", "")
	$s = StringReplace($s, "'", "")
	$s = StringReplace($s, "&", "and")
EndFunc ; fixString

; Description: remove spaces commas of a string
; Parameters:
; 		$n - the string to manipulate
Func fixStringNumber(ByRef $n)
	$n = StringStripWS($n, 8)
	$n = StringStripCR($n)
	$n = StringReplace($n, ",", "")
EndFunc


; Description: change the current position of the mouse
; Parameters:
; 		$x, $y - Mouse location to add or subtract
; 		$isAdd - Values must only be 0 OR 1, 0 = Move Right; 1 = Move Left
Func mMoveChange($x, $y, $isAdd)
	$mPos = MouseGetPos()
	$mSpeed = 5
	If $isAdd = 0 Then
		MouseMove($mPos[0] + $x, $mPos[1] + $y, $mSpeed)
	ElseIf $isAdd = 1 Then
		MouseMove($mPos[0] - $x, $mPos[1] - $y, $mSpeed)
	EndIf
EndFunc ; mMoveChange

; Description: searches the image and moves mouse
; Parameters:
; 		$pic - The path of the picture to search
;		$b - If True the mouse will move to the $x and $y
; Return Value:
; 		True - Success otherwise False
#cs
Func imgSearch($pic, $b = True)

	$x = 0
	$y = 0
	$isFound = False
	#cs
	do
		$result = _ImageSearch($pic, 1, $x, $y, 0, 0)
		ConsoleWrite($result)
	until $result = 1;
	#ce
	;MsgBox(0, "", "will now search")
	for $i = 1 To 20
		$isFound = _ImageSearch($pic, 1, $x, $y, 0, 0)
		Sleep(10)
		ConsoleWrite($isFound)
	Next
	MsgBox(0, @WorkingDir & "/" & $pic, $isFound)

	If $isFound = 1 Then
		If $b Then
			MouseMove($x, $y, 10)
		EndIf
		Return True
	Else
		Return False
	EndIf

EndFunc ; imgSearch
#ce

; Description: To activate a window If it already exists
; Parameters:
; 		$winTitle - Title of the window to activate
; 		$winDesc - Description to output when the window cannot be activated (opt def:"BIR")
Func activateWin(Const $winTitle, Const $winDesc = "BIR")

	Local $isActive = WinActivate($winTitle, "")
	If $isActive = 0 Then
		_BlockInputEx(0)
		MsgBox(16, "Not Opened", "The "& $winDesc &" window is not opened!")
		exitz()
	EndIf

EndFunc; activateWin

; Description: To exit/stop the script
Func exitz()
	; unblock keys
	_BlockInputEx(0)

	$isContinue = False
	MsgBox(0, "Exit", "Script will now exit...")
	Exit 0
EndFunc ; exitz

; END OF FUNCTIONS


; EVENTS

; Description: displays File Open Dialog to select the excel file
Func browse()
	$excelFile = FileOpenDialog("Select Excel File", @DocumentsCommonDir, "Excel(*.xls; *.xlsx)", $FD_FILEMUSTEXIST  + $FD_PATHMUSTEXIST, "", $frmMain)
	If @error = 1 Then
		Return
	EndIf
	MsgBox(0, "", $excelFile)
	;$excelFile = StringRegExp($excelFile, ".*\\(.*)[^(.xlsx?)]", 1)
	$excelFile = StringRegExp($excelFile, ".*\\(.+)\.(.+)", 1)

	GUICtrlSetData($txtFileName, $excelFile[0])
	FileChangeDir(@ScriptDir)
EndFunc ; browse


; Description: enables/disables the 3 radio buttons on Services
; Parameters:
; 		$b = If True enable, else disable button
Func enableRadioServices($b)
	Local $val = 64
	If $b = False Then
		$val *= 2
	EndIf

	GUICtrlSetState($radServices, $val)
	GUICtrlSetState($radCapitalGoods, $val)
	GUICtrlSetState($radOtherGoods, $val)
EndFunc ; enableRadioServices

; Description: sets the modes and excel title and calls the automate func
Func start()

	; #cs
	; read excel title
	$ExcelTitle = GUICtrlRead($txtFileName)
	$ExcelTitle = StringStripCR($ExcelTitle)
	$ExcelTitle = StringStripWS($ExcelTitle, 3)

	If StringLen($ExcelTitle) = 0 Then
		MsgBox(48, "No Input", "Please enter a file name of the Excel", 0, $frmMain)
		Return
	EndIf

	; note
	MsgBox(64, "Note", "Note:" & @CRLF & "* You can stop the script by pressing the END on your keyboard." & @CRLF _
		& "* The mouse and keyboard are disabled so the script will not be interfered, if you wish to stop the script just press the END key.")

	; check if windows are opened
	If WinExists($ExcelTitle) = 0 Then
		MsgBox(16, "Error", "The '" & $ExcelTitle & "' Excel window is not opened.", 0, $frmMain)
		Return
	elseif WinExists($BIRtitle) = 0 Then
		MsgBox(16, "Error", "The " & $BIRtitle & " window is not opened.", 0, $frmMain)
		Return
	EndIf
	; #ce

	; block now
	_BlockInputEx(1, "{END}")

	; check script mode
	activateWin($BIRtitle)
	Local $modeSearch = False
	Local $isValidMode = False
	Local $modeString = ""


	Switch $scriptMode
		Case $M_SALES
			$modeSearch = getMode($M_SALES)
			$modeString = "Sales"
		Case $M_PURCHASE
			$modeSearch = getMode($M_PURCHASE)
			$modeString = "Purchase"

			; set purchase mode
			If GUICtrlRead($radServices) = $GUI_CHECKED Then
				$purchaseMode = $P_SERVICES
			ElseIf GUICtrlRead($radCapitalGoods) = $GUI_CHECKED Then
				$purchaseMode = $P_CAPITAL_GOODS
			ElseIf GUICtrlRead($radOtherGoods) = $GUI_CHECKED Then
				$purchaseMode = $P_OTHER_GOODS
			EndIf
	EndSwitch

	If $modeSearch = False Then
		While Not $isValidMode
			_BlockInputEx(0)
			Local $opt = MsgBox(48+1, "Warning", "The BIR is not on " & $modeString &  @CRLF & "Please go to " & $modeString & " in BIR then press 'OK'", 0, $frmMain)
			If $opt = 1 Then
				activateWin($BIRtitle)
				Switch $scriptMode
					Case $M_SALES
						$isValidMode = getMode($M_SALES)
					Case $M_PURCHASE
						$isValidMode = getMode($M_PURCHASE)
				EndSwitch
				If $isValidMode = True Then
					MsgBox(0, "", "Found!" & @CRLF & "Script will now continue...", 3, $frmMain)
					_BlockInputEx(1, "{END}")
				EndIf
			Else
				$isValidMode = False
				Return
			EndIf
		WEnd
	EndIf

	GUIDelete($frmMain)

	; run automation
	automate()

EndFunc ; exec

; Description: To close the main GUI
Func mainGUIClose()

	$opt = MsgBox(32+4, "Close", "Are you sure you want to exit?", 0, $frmMain)
	If $opt = 6 Then ; exit if YES(6)
		GUIDelete($frmMain)
		Exit 0
	EndIf

EndFunc ; mainGUIClose

; Description: About Item Menu
Func itmAbout()
	MsgBox(64, "BIR SCRIPT", "Script for automatic encoding of Sales and Purchases in BIR Relief System" & @CRLF & "Version: 1.2.2" & @CRLF & _
	"Date Created: May 6, 2016" & @CRLF & "Author: Zaerald Lungos" & @CRLF & @CRLF & "The icon used was downloaded from http://www.iconarchive.com/tag/script", 0, $frmMain)
EndFunc ; itmAbout

; END OF EVENTS
