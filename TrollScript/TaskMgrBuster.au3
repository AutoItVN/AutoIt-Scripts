HotKeySet("{END}", "exitNow")


While True
	Sleep(50)
	;$process = ProcessExists("z")
	$process = ProcessExists("taskmgr.exe")
	If Not $process = 0 Then
		ProcessClose($process)
	EndIf
WEnd

; FUNCTIONS

Func exitNow()
	Exit 0
EndFunc

; END FUNCTIONS