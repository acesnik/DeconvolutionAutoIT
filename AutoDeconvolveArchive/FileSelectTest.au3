$FileCt = 97
For $i = 0 To $FileCt - 2 ;Note: To is inclusive

		 MsgBox(0, "AutoDeconvolve", "Processing File #" & $i + 1, 0.5)
		 ;click on the second raw file, go down to the file to process
		 If $i < 40 Then
			MouseClick("primary", 738 , 340, 1)
			For $j = 1 To $i
			   Send("{DOWN}", 0)
			   sleep(10)
			Next
			If $i = 0 Then
			   $i = 33
			EndIf
		 Else
			MouseClick("primary", 735 , 844, 1)
			Send("{DOWN}", 0)
			sleep(10)
		 EndIf

		 MouseClick("primary", 1129, 895, 1) ;click load
		 sleep(10)

		 ;get out of the check window called "Protein Deconvolution"
		 ;WinWaitActive("Protein Deconvolution")
		 ;Send("{LEFT}{ENTER}",0)
		 ;WinWaitActive("Thermo Protein Deconvolution")
		 ;MouseClick("primary", 1337, 285, 1) ;click top method
		 ;MouseClick("primary", 1700, 901, 1) ;click load method button
		 ;WinWaitActive("Thermo Protein Deconvolution")

Next