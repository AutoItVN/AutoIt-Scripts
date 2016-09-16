#include <ImageSearch2015.au3>

$x = 0
$y = 0

$picture = "Serts.bmp"
Sleep(500)
do
	$result = _ImageSearch($picture, 1, $x, $y, 0, 0)
	ConsoleWrite($result)
until $result = 1;

if $result = 1 Then
	MouseMove($x, $y, 10)
Else
	MsgBox(0, "Nothing Found!", "No Results Found!")
EndIf
