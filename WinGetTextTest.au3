sleep(5000)
$text = WinGetText("[ACTIVE]")
MsgBox(0, "All Text", $text)
Local $arr = StringSplit($text, @LF)
MsgBox(0,"Just common string", $arr[5])
If $arr[5] == "Do you want to replace it?" Then
   MsgBox(0, "Success", "success")
EndIf
