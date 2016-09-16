#include<File.au3>

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=DOTA 2.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
HotKeySet("{NUMPADADD}", "playStop")
HotKeySet("{NUMPADSUB}", "exitz")
HotKeySet("{NUMPAD0}", "play0")
HotKeySet("{NUMPAD1}", "play1")
HotKeySet("{NUMPAD2}", "play2")
HotKeySet("{NUMPAD3}", "play3")
HotKeySet("{NUMPAD4}", "play4")
HotKeySet("{NUMPAD5}", "play5")
HotKeySet("{NUMPAD6}", "play6")
HotKeySet("{NUMPAD7}", "play7")
HotKeySet("{NUMPAD8}", "play8")
HotKeySet("{NUMPAD9}", "play9")


; read file
Local $file = "sounds.txt"
Local $sounds[10];

for $i = 1 to 10
	$sounds[$i-1] = FileReadLine($file, $i)
Next

MsgBox(0, "Started...", "Program Started!")

While True
	Sleep(100)
WEnd

Func play0()
	SoundPlay($sounds[0] & ".mp3")
EndFunc

Func play1()
	SoundPlay($sounds[1] & ".mp3")
EndFunc

Func play2()
	SoundPlay($sounds[2] & ".mp3")
EndFunc

Func play3()
	SoundPlay($sounds[3] & ".mp3")
EndFunc

Func play4()
	SoundPlay($sounds[4] & ".mp3")
EndFunc

Func play5()
	SoundPlay($sounds[5] & ".mp3")
EndFunc

Func play6()
	SoundPlay($sounds[6] & ".mp3")
EndFunc

Func play7()
	SoundPlay($sounds[7] & ".mp3")
EndFunc

Func play8()
	SoundPlay($sounds[8] & ".mp3")
EndFunc

Func play9()
	SoundPlay($sounds[9] & ".mp3")
EndFunc

Func playStop()
	SoundPlay("")
EndFunc

Func exitz()
	MsgBox(0, "Exit", "Exitting now...")
	Exit 0
EndFunc