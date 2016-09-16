#include <ImageSearch2015.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

HotKeySet("{BREAK}", "exitz")

AutoItSetOption("SendKeyDelay", 100) ; 100
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("GUIOnEventMode", 1)

; CONSTANTS
; WinTitles
Local Const $BIRtitle = "ATTACHMENTS DATA ENTRY SYSTEM"
Local $ExcelTitle = ""

; purchase modes
Local Const $P_SERVICES = 0
Local Const $P_CAPITAL_GOODS = 1
Local Const $P_OTHER_GOODS = 3

Local $isContinue = True
Local $purchaseMode = -1
Local $isTest = False

; BIR Dialogs
; BIR System Message : [CLASS:#32770] - Name err
; Microsoft Visual FoxPro: [CLASS:#32770] - Exceed
; BIR System Message

; SCRIPT START

; MAINGUI
$frmMain = GUICreate("BIR System Script", 432, 114, 432, 221)
$itmAbout = GUICtrlCreateMenu("&About")
GUISetBkColor(0xFFFFFF)
$lblExcel = GUICtrlCreateLabel("Excel File Name:", 21, 16, 83, 17)
$txtFileName = GUICtrlCreateInput("", 112, 12, 305, 21)
$btnStart = GUICtrlCreateButton("Start!", 256, 49, 75, 25)
GUICtrlSetOnEvent($btnStart, "exec")
$btnCancel = GUICtrlCreateButton("Cancel", 344, 49, 75, 25)
GUICtrlSetOnEvent($btnCancel, "mainGUIClose")
$chkTest = GUICtrlCreateCheckbox("Test", 16, 48, 49, 17)
GUISetState(@SW_SHOW)

; states
GUISetState(@SW_SHOW)
; END MAINGUI

; keep GUI visible
While 1
	Sleep(20)
WEnd

; Description: the main automation function
Func automate()
	; go to excel
	activateWin($ExcelTitle, "Excel")

	; set fill color to yellow
	Send("!h{ESC 2}")
	Sleep(100)
	imgSearch("res/fill.bmp")
	mMoveChange(15, 0, 0)
	MouseClick("left")
	Sleep(100)
	imgSearch("res/fillclr.bmp")
	MouseClick("left")
	Send("^z")

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
		if StringLen($TIN) = 0 Then
			exitz()
		EndIf

		; continue copy TIN
		$LTIN = StringLeft($TIN, 11)
		$RTIN = "0" & StringRight($TIN, 3)
		activateWin($BIRtitle)
		Send("!r")
		; check if QUESTION Pops
		$bExists = WinExists("QUESTION")
		if $bExists = 1 Then
			Send("{SPACE}")
		EndIf

		Send("!a")
		ClipPut($LTIN)
		Send("^v{TAB}")
		ClipPut($RTIN)
		Send("^v{TAB}")

		Sleep(30)
		; get caret position to check if TIN already exists
		Local $caretPos = WinGetCaretPos()
		; MsgBox(0, "", $caretPos[0] & ", " & $caretPos[1])

		if $caretPos[1] >= 260 Then
			$isTINexists = True
		EndIf

		; copy reg name and address
		if $isTINexists = False Then
			Local $copied = ""
			; copy reg name
			; 50 chr limit, '.
			activateWin($ExcelTitle, "Excel")
			Send("{HOME}{RIGHT}^c")
			activateWin($BIRtitle)
			Send("^v{TAB 4}")
		EndIf


		; copy code and amount
		activateWin($ExcelTitle, "Excel")
		Send("{HOME}{RIGHT 2}^c")
		activateWin($BIRtitle)
		Send("^v{TAB}")
		activateWin($ExcelTitle, "Excel")
		Send("{RIGHT}^c")
		activateWin($BIRtitle)
		Send("^v{TAB 2}")

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
			$msg = MsgBox(4, "Not Equal", "Total input Tax is not equal to Sum of TAX" & @CRLF & "BIR Total:        " & $BIRtotal & _
			@CRLF & "In Excel Total: " & $ExcelTotal & @CRLF & @CRLF & "Do you still want to  proceed?" )
			if $msg = 7 Then
				exitz()
			EndIf
		EndIf

		; fill color finished
		Send("{ESC}{HOME}^+{RIGHT}")
		Sleep(100)
		imgSearch("res/fill.bmp")
		Sleep(50)
		MouseClick("left")
		Send("{ESC}{DOWN}")


		; Save
		if Not $isTest Then
			MsgBox(0, "", "Will save!")
			Send("!s")
			WinWaitActive("QUESTION", "", 3)
			Send("!y")
		EndIf

		; continue?
		$opt = MsgBox(4, "Continue?", "Want to continue?")
		if $opt = 7 Then
			exitz()
		EndIf


	WEnd
EndFunc ; automate


; SCRIPT END

; FUNCTIONS

; Description: check if the string exceeds the given limit
; Parameters:
; 		$s - String to check
; 		$l - the number of limit
Func chkStringLimit($s, $l)
	; count string
	Local $count = StringLen($s)
	if $count >= $l Then
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
	if $isAdd = 0 Then
		MouseMove($mPos[0] + $x, $mPos[1] + $y, $mSpeed)
	ElseIf $isAdd = 1 Then
		MouseMove($mPos[0] - $x, $mPos[1] - $y, $mSpeed)
	EndIf
EndFunc ; mMoveChange

; Description: the image to search
; Parameters:
; 		$pic - The path of the picture to search
; Return Value:
; 		True - Success otherwise False
Func imgSearch($pic)

	$x = 0
	$y = 0

	do
		$result = _ImageSearch($pic, 1, $x, $y, 0, 0)
		ConsoleWrite($result)
	until $result = 1;

	if $result = 1 Then
		MouseMove($x, $y, 10)
		Return True
	Else
		Return False
	EndIf

EndFunc ; imgSearch

; Description: To actiavate a window if it already exists
; Parameters:
; 		$winTitle - Title of the window to activate
; 		$winDesc - Description to output when the window cannot be activated (opt def:"BIR")
Func activateWin(Const $winTitle, Const $winDesc = "BIR")

	Local $isActive = WinActivate($winTitle, "")
	if $isActive = 0 Then
		MsgBox(16, "Not Opened", "The "& $winDesc &" window is not opened!")
		exitz()

	EndIf

EndFunc; activateWin

; Description: To exit/stop the script
Func exitz()
	$isContinue = False
	MsgBox(0, "Exit", "Script will now exit...")
	Exit 0
EndFunc ; exitz

; END OF FUNCTIONS

; #######
; EVENTS

; Description: sets the modes and excel title and calls the automate func
Func exec()

	; read excel title
	$ExcelTitle = GUICtrlRead($txtFileName)
	$ExcelTitle = StringStripCR($ExcelTitle)
	$ExcelTitle = StringStripWS($ExcelTitle, 3)

	; check if there's input
	if StringLen($ExcelTitle) = 0 Then
		MsgBox(48, "No Input", "Please enter a file name of the Excel")
		Return
	EndIf

	; check if windows are opened
	if WinExists($ExcelTitle) = 0 Then
		MsgBox(16, "Error", "The '" & $ExcelTitle & "' Excel window is not opened.")
		Return
	elseif WinExists($BIRtitle) = 0 Then
		MsgBox(16, "Error", "The " & $BIRtitle & " window is not opened.")
		Return
	EndIf

	; check if a test
	if GUICtrlRead($chkTest) = $GUI_CHECKED Then
		$isTest = True
	EndIf

	GUIDelete($frmMain)
	automate()
EndFunc ; exec

; Description: To close the main GUI
Func mainGUIClose()

	$opt = MsgBox(32+4, "Close", "Are you sure you want to exit?")
	if $opt = 6 Then ; exit if YES(6)
		GUIDelete($frmMain)
		Exit 0
	EndIf

EndFunc ; mainGUIClose

; Description: About Item Menu
Func itmAbout()
	MsgBox(64, "BIR SCRIPT", "Script for automatic endcoding of Purchases in BIR System" & @CRLF & "Version: 1.0" & @CRLF & _
	"Date Created: May 6, 2016" & @CRLF & "Author: Zaerald Lungos")
EndFunc ; itmAbout

; END OF EVENTS
; #############