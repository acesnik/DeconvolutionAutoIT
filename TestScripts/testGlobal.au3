 #include <Constants.au3> ; only required for MsgBox constants

 Global $Var_1 = 3
 Global $Var_2 = 1

 ; Read the variables
 MsgBox($MB_ICONINFORMATION + $MB_TOPMOST, "Reading", "In The Main Script" & @CRLF & @CRLF & "Variable 1 = " & $Var_1 & @CRLF & "Variable 2 = " & $Var_2)

 ; Read the variables in a function
 _Function()

 ; And now read the variables again - see that $iVar_2 has changed
 MsgBox($MB_ICONINFORMATION + $MB_TOPMOST, "Reading", "Back In The Main Script" & @CRLF & @CRLF & "Variable 1 = " & $Var_1 & @CRLF & "Variable 2 = " & $Var_2)

 Func _Function()
 	; Read the variables
 	MsgBox($MB_ICONINFORMATION + $MB_TOPMOST, "Reading", "In The Function" & @CRLF & @CRLF & "Variable 1 = " & $Var_1 & @CRLF & "Variable 2 = " & $Var_2)
 	; Now let us change one of the variables WITHIN the function
 	$Var_2 = "Changed Variable 2"
 EndFunc