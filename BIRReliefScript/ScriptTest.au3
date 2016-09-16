#include <ImageSearch2015.au3>

imgSearch("res/purchases.png")

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