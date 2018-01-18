;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
; File   :  AutoAvgQualBrowser
; Author :  Anthony J. Cesnik
; History:  141024 v0.1 Coded to initial functionality
;			141024 v0.2	Changed mouseclick position for exporting spectra
;						b/c it gives a different menu if right click on a peak label.
;			141105 v0.3 Modify to have #of spectra to average as a userInput
;						and for mixed file types where not all spectral scans are averaged.
;			141110 v0.4 Discrete things to do for averaging 1, 2, or >2 spectra.
;
; Credit :  Brian Frey did the heavy lifting of figuring out the best workflow
;           in Protein Deconvolution/Qual Browser for current purposes!
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

$SaveLocation = InputBox("AutoAvgQualBrowser", "Enter the folder path in which you would like to save these averaged spectra." & @LF & @LF & "Leave this empty if you would like to save in the C:\Xcalibur\Data default directory.")
$BaseName = InputBox("AutoAvgQualBrowser", "Enter the base file name you would like to use." & @LF & @LF & "Example: Base file name, File, will produce File_Avg1-3 for average spectra 1-3.")
$ScanInterval= InputBox("AutoAvgQualBrowser", "Enter the counting interval of spectra." & @LF & @LF & "e.g. enter 4, if full/SIM scans occur every 4th scan.")
If $ScanInterval <> "" Then
   $NumberSpectraToAvg = InputBox("AutoAvgQualBrowser", "Enter the number of spectra you would like averaged together.")
   If $NumberSpectraToAvg <> "" Then
	  $FinalSpect = InputBox("AutoAvgQualBrowser", "Enter the final spectrum you would like to process." & @LF & @LF & "Hit OK, then you will be asked to select the beginning spectra.")
	  If $FinalSpect <> "" Then
		 $StartSpect = InputBox("AutoAvgQualBrowser", "Now, select the first average spectrum you would like to process." & @LF & @LF & "Enter the starting spectrum number.")
		 If $StartSpect <> "" Then

			;Main loop: continue until the right-most spectrum in this range hits the final spectrum and finishes processing it
			For $i = $StartSpect to $FinalSpect - (($NumberSpectraToAvg -1)*$ScanInterval) Step $ScanInterval ;$i is the start spectrum in the current range

			   ;Select the next set of spectra
			   If $i <> $StartSpect Then
				  Send("{RIGHT}", 0)
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
			   Send($BaseName)
			   If $NumberSpectraToAvg == 1 Then
				  Send("_" & $i & "{ENTER}", 0) ;enter spectrum numbers
			   Else
				  If $NumberSpectraToAvg == 2 Then
					 Send("_Avg" & $i & "," & ($i + $ScanInterval) & "{ENTER}", 0) ;enter spectrum numbers
				  Else
				  Send("_Avg" & $i & "-" & ($i + (($NumberSpectraToAvg -1)*$ScanInterval)) & "{ENTER}", 0) ;enter spectrum numbers
				  EndIf
			   EndIf
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
	  MsgBox(0, "AutoAvgQualBrowser", "You didn't enter the number of spectra to average."
   EndIf
Else
   MsgBox(0, "AutoAvgQualBrowser", "You didn't enter the scan counting interval."
EndIf