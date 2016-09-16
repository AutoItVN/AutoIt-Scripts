#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=steam.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("SendKeyDelay", 70)


Local $isRunActive = False
Local $iChrome = 1
Local $asSites = ["messenger.com", "fb.me"]

BlockInput(1)

MsgBox(0, "My Favorite application.", "You will have a phenomenal experience." & @CRLF & "Please enjoy!!! :D", 3)

$isRunActive = openRun()

; open chrome
if $isRunActive Then
	Send("chrome{ENTER}")
	$iChrome = WinWaitActive("New Tab - Google Chrome", "", 3)
EndIf

if $iChrome = 0 Then
	; open internet explorer
	$isRunActive = openRun()
	if $isRunActive Then
		Send("iexplore{ENTER}")
	EndIf
EndIf

; maximize
$iChrome = WinSetState("Google Chrome", "" , @SW_MAXIMIZE)
if $iChrome = 0 Then
	WinSetState("[CLASS:IEFrame]", "" , @SW_MAXIMIZE)
EndIf

; browse now
for $i in $asSites
	Send($i & "{ENTER}^t")
Next

Send("^w") ; close last tab

; navigate
AutoItSetOption("SendKeyDelay", 3000) ; lower delay for slow navigation
for $i = 1 to UBound($asSites) * 2
	Send("^{TAB}")
Next
AutoItSetOption("SendKeyDelay", 70) ; return to normal delay

MouseMove(1341, 10)
BlockInput(0)
MsgBox(0, "Finished", "So you had a good experience??? HAHAHAAH!!!")


; functions
Func openRun()
	$iRun = 1
	Send("{LWINDOWN}r{LWINUP}")
	$iRun = WinWaitActive("Run", "", 5)
	if $iRun = 0 Then
		exitNow()
	EndIf
	Return True
EndFunc

Func exitNow()
	BlockInput(0)
	Exit 0
EndFunc






