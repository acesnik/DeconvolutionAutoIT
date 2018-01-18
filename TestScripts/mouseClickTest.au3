sleep(5000)
$aPos = MouseGetPos()
MsgBox(0, "Mouse x, y:", $aPos[0] & ", " & $aPos[1])
;WinWait("Thermo Protein Deconvolution")
sleep(1000)
MouseClick("primary", $aPos[0] , $aPos[1], 1) ;click on the second raw file
