#include <WinAPI.au3>

HotKeySet("{END}", "_UnBlock")

Global $key = DllCallbackRegister("_KeyProc", "int", "int;ptr;ptr")
Global $mouse = DllCallbackRegister ("_Mouse_Handler", "int", "int;ptr;ptr")

MsgBox(0, "", 'Now we block mouse and all keyboard keys except a "F3"')

Global $hookKey = _WinAPI_SetWindowsHookEx($WH_KEYBOARD_LL, DllCallbackGetPtr($key), _WinAPI_GetModuleHandle(0), 0)
Global $hookMouse = _WinAPI_SetWindowsHookEx($WH_MOUSE_LL, DllCallbackGetPtr($mouse), _WinAPI_GetModuleHandle(0), 0)

While 1
    Sleep(100)
WEnd

Func _UnBlock()
    DllCallbackFree($key)
    DllCallbackFree($mouse)
    _WinAPI_UnhookWindowsHookEx($hookKey)
    _WinAPI_UnhookWindowsHookEx($hookMouse)
    MsgBox(0, "_UnBlock", "Input unblocked")
    Exit
EndFunc

Func _KeyProc($nCode, $wParam, $lParam)
    If $nCode < 0 Then Return _WinAPI_CallNextHookEx($hookKey, $nCode, $wParam, $lParam)

    Local $KBDLLHOOKSTRUCT = DllStructCreate("dword vkCode;dword scanCode;dword flags;dword time;ptr dwExtraInfo", $lParam)
    Local $vkCode = DllStructGetData($KBDLLHOOKSTRUCT, "vkCode")

    If $vkCode <> 0x23 Then Return 1

    _WinAPI_CallNextHookEx($hookKey, $nCode, $wParam, $lParam)
EndFunc

Func _Mouse_Handler($nCode, $wParam, $lParam)
    If $nCode < 0 Then Return _WinAPI_CallNextHookEx($hookMouse, $nCode, $wParam, $lParam)

    Return 1
EndFunc