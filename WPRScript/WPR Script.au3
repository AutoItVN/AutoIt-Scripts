#cs

WPR Script

April 26, 2016

Author: Zaerald Lungos
#ce
#include <ImageSearch2015.au3>


HotKeySet("{BREAK}", "exitz")
AutoItSetOption ( "SendKeyDelay", 100)



; vars
; Powerpoint Title
Local $sPowerpointTitle = "SSS WPR Apr 18-24 2016 Edralin.pptx - PowerPoint"
Local $sWindowDate = "April 18 to 24"
Local $aDates = ["April 18", "April 19", "April 20", "April 21", "April 22", "April 23", "April 24"]

; Comparative Title
Local $sComparativeTitle = "ctg sales monitoring April 18-24.xls - Excel"

windowActivate($sPowerpointTitle)

;resize
WinSetState($sPowerpointTitle, "", @SW_MAXIMIZE)
imgSearch("res/zmdlg.bmp")
MouseClick("left")
Send("!p70+5{ENTER}")

; go to page 1
Send("{HOME}")

; page 1
Send("{END}+{HOME}+^{RIGHT}")
Send($sWindowDate)

; get data from comparative
windowActivate($sComparativeTitle)
Send("^{HOME 2}")
Sleep(100)
Send("^fEdralin{ENTER 2}")
Send("{ESC}{DOWN}{RIGHT}")
Send("^+{DOWN}^+{RIGHT}")
Send("^c")

; go back to ppt
windowActivate($sPowerpointTitle)

; go to page 3
Send("{PGDN 2}")

Sleep(10)
; page 3
; MouseClick("left", 832, 480, 1, 5)
imgSearch("res/tblpg3.bmp")
exit 0
Send("^{HOME 3}^{DOWN}{DOWN}^v")
Send("+{F10}{DOWN 3}{ENTER}")
WinWaitActive("Font", "")
Send("!S14{ENTER}")

; go to page 7
Send("{PGDN 4}")

; page 7
MouseClick("left", 767, 475, 1, 5)
Send("{F2}{HOME}^{HOME 2}")


Func mMoveChange($x, $y, $isAdd)
	$mPos = MouseGetPos()
	$mSpeed = 5
	if $isAdd = 0 Then
		MouseMove($mPos[0] + $x, $mPos[1] + $y, $mSpeed)
	ElseIf $isAdd = 1 Then
		MouseMove($mPos[0] - $x, $mPos[1] - $y, $mSpeed)
	EndIf
EndFunc


Func imgSearch($pic)

	$x = 0
	$y = 0

	do
		$result = _ImageSearch($pic, 1, $x, $y, 0, 0)
		ConsoleWrite($result)
	until $result = 1;

	if $result = 1 Then
		MouseMove($x, $y, 10)
	Else
		MsgBox(0, "Nothing Found!", "No Results Found!")
	EndIf

EndFunc

Func windowActivate($windowTitle)

	; activate window
	Local $bIsWindowActive = WinActivate($windowTitle)

	;check if exists
	If $bIsWindowActive = False Then
		MsgBox(0, "Error", "The Window is not opened!")
		Exit 0
	EndIf
EndFunc


Func exitz()
	MsgBox(0, "Exit", "Script will now exit...")
	Exit 0
EndFunc











