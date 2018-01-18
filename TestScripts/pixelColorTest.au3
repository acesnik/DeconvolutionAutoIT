$aPos = 0
$xPos = InputBox("Choose a position?", "Choose an x-position you would like to check the color for." & @LF & "Leave empty and hover over position to check the color for a new position.")
If $xPos = "" Then
   sleep(5000)
   $aPos = MouseGetPos()
   MsgBox(0, "Color at x, y:", "Color at x, y: " & $aPos[0] & ", " & $aPos[1] & " is " & PixelGetColor($aPos[0], $aPos[1]))
   sleep(1000)
   MouseClick("primary", $aPos[0] , $aPos[1], 1) ;click on the second raw file
Else
   $yPos = InputBox("Choose a position?", "Choose a y-position you would like to check the color for.")
   $x = Number($xPos)
   $y = Number($yPos)
   MsgBox(0, "Color at x, y:", "Color at x, y: " & $x & ", " & $y & " is " & PixelGetColor($x, $y))
   sleep(1000)
   MouseClick("primary", $x , $y, 1) ;click on the second raw file
EndIf

