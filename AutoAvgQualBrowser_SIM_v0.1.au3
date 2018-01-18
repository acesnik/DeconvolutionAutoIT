;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
; File   :  AutoAvgQualBrowser
; Author :  Anthony J. Cesnik & Brian Frey
; History:  141112 v0.1 Started with AutoAvgQualBrowser_v0.4 and modified to move
;						right arrow key the ScanInterval # of times. Note that
;						no averaging is done here b/c consecutive SIM scans are not same m/z range.
;
; Credit :  Brian Frey did the heavy lifting of figuring out the best workflow
;           in Protein Deconvolution/Qual Browser for current purposes!
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

$SaveLocation = InputBox("AutoAvgQualBrowser", "Enter the folder path in which you would like to save these averaged spectra." & @LF & @LF & "Leave this empty if you would like to save in the C:\Xcalibur\Data default directory.")
$BaseName = InputBox("AutoAvgQualBrowser", "Enter the base file name you would like to use." & @LF & @LF & "Example: Base file name, File, will produce File_SIM_22 for SIM spectrum 22.")
$ScanInterval= InputBox("AutoAvgQualBrowser", "Enter the counting interval of spectra." & @LF & @LF & "e.g. enter 4, if full/SIM scans occur every 4th scan.")
If $ScanInterval <> "" Then
	  $FinalSpect = InputBox("AutoAvgQualBrowser", "Enter the final spectrum you would like to process." & @LF & @LF & "Hit OK, then you will be asked to select the beginning spectra.")
	  If $FinalSpect <> "" Then
		 $StartSpect = InputBox("AutoAvgQualBrowser", "Now, select the first spectrum you would like to process." & @LF & @LF & "Enter the starting spectrum number.")
		 If $StartSpect <> "" Then

			;Main loop: continue until the right-most spectrum in this range hits the final spectrum and finishes processing it
			For $i = $StartSpect to $FinalSpect  Step $ScanInterval ;$i is the start spectrum in the current range

			   ;Select the next set of spectra
			   If $i <> $StartSpect Then
				  For $j = 1 to $ScanInterval ; loop to hit right arrow key correct # of times to get to next SIM spectrum
					 Send("{RIGHT}", 0)
					 Sleep(100)
				  Next
				  Do
					 sleep(100)
					 $color = PixelGetColor(1059,754)
				  Until $color <> 16777215 ;wait until pixel in the spectrum window is not white
			   EndIf

			   ;Export the averaged spectra
			   MouseClick("secondary", 1580,596) ;Right click spectrum
			   MouseClick("primary", 1638, 672) ;Click export
			   sleep(200)
			   MouseClick("primary", 1745,707) ;Click Write to Raw File

			   ;Save As window
			   WinWaitActive("Save As")
			   sleep(10)
			   MouseClick("primary",1189,280) ;enter the directory name
			   Send($SaveLocation)
			   Send("{ENTER}", 0)
			   sleep(200) ;needs this fairly long wait
			   Send("{TAB 6}{BS}", 0) ;get back to the file name fields
			   Send($BaseName & "_SIM_" & $i & "{ENTER}", 0);enter spectrum numbers
			   sleep(500) ;wait a bit
			   While MouseGetCursor() = 15 ;wait until hour-glass cursor is gone
				  sleep(100)
			   WEnd
			   WinWaitActive("Thermo Xcalibur Qual Browser") ;wait until the Qual Browser window is active again

		   Next ;End of main loop

		 Else
			MsgBox(0, "AutoAvgQualBrowser", "You didn't enter a start spectrum number."
		 EndIf
	  Else
		 MsgBox(0, "AutoAvgQualBrowser", "You didn't enter a final spectrum number."
	  EndIf
Else
   MsgBox(0, "AutoAvgQualBrowser", "You didn't enter the scan counting interval."
EndIf