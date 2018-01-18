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
;           141119 v0.5 Adding functionality to compensate for the Extract algorithm
;                       bug that prevents it from dumping the memory after writing
;                       to a RAW file after averaging the spectra.
; Instead of having the user select the first three files, and then let the
; program do the right-button shuffle, this program uses the "Range" function
; to select the correct three spectra. Also, it adds the functionality of
; progressing through a list of runs (~30 are expected in this case).
;           141120      Tested on a second file, and it is not able to deal with
;                       the different spectrum widths, so I'm going to acquire it dynamically.
;
; At least for now, I'm going to get omit the counting interval of the spectra
; and only make the thing work for three spectra
;
; Credit :  Brian Frey did the heavy lifting of figuring out the best workflow
;           in Protein Deconvolution/Qual Browser for current purposes!
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;This doesn't work to close out windows when you hit close right now...
;Not a big deal though
Func CatchErrors()
   If @error Then
	  Exit
   EndIf
EndFunc

;$SaveLocation = InputBox("AutoAvgQualBrowser", "Enter the folder path in which you would like to save these averaged spectra." & @LF & @LF & "Leave this empty if you would like to save in the C:\Xcalibur\Data default directory.")
;$BaseName = InputBox("AutoAvgQualBrowser", "Enter the base file name you would like to use." & @LF & @LF & "Example: Base file name, File, will produce File_Avg1-3 for average spectra 1-3.")
;$FolderName = ""
;While $FolderName == ""
;   $FolderName = InputBox("Initialize AutoAvgQualBrowser", "What is the name of the folder with the raw files you would like to process? (Not the path.)")
;   CatchErrors()
;WEnd
$NumFiles = ""
While $NumFiles == ""
   $NumFiles = InputBox("Initialize AutoAvgQualBrowser", "Make a folder with only the raw files you would like to process. How many files are there?")
   CatchErrors()
WEnd
$qualBrowsing = ProcessWait("QualBrowser.exe", 1)
If $qualBrowsing Then
   MsgBox(0, "Initialize AutoAvgQualBrowser", "Close Qual Browser, please. But keep that folder in mind!")
   CatchErrors()
EndIf
MsgBox(0, "Initialize AutoAvgQualBrowser", "Please open that folder in Windows Explorer, select the first file, and then hit okay.")
CatchErrors()

sleep(500)
Send("{ENTER}",0) ;Opens the first folder in qual browser
$opened = WinWaitActive("Thermo Xcalibur Qual Browser", "", 2) ;Wait for the window to open.
;$alreadyRunning = UBound(ProcessList("AutoIt3.exe")) > 1
;If $alreadyRunning Then
   ;MsgBox(0, "Initialize AutoAvgQualBrowser", "An AutoIt script is already running. This program doesn't work well when another of the same script is running.")
   ;CatchErrors()
If Not $opened Then
   Send("{ENTER}",0) ;Opens the first folder in qual browser
   $opened = WinWaitActive("Thermo Xcalibur Qual Browser", "", 10) ;Wait for the window to open. If it doesn't call an error message.
   If Not $opened Then
	  MsgBox(0, "Initialize AutoAvgQualBrowser", "Qual Browser appears to have failed to open. Please check for an issue.")
	  CatchErrors()
   EndIf
Else

   ;Select the first retention time range.
   MouseClick("secondary",1116,296)
   sleep(10)
   MouseClick("primary",1163,444)
   WinWaitActive("Chromatogram Ranges")
   $firstStartT = 39
   $firstEndT = 40
   sleep(100)
   Send($firstStartT & "-" & $firstEndT)
   Send("{ENTER}",0)
   Do
	  sleep(100)
	  $color = PixelGetColor(1605,167)
   Until $color == 15921906 ;wait until pixel in the spectrum window is grey
   MouseClick("primary", 1904,593) ;Click the spectrum tack button

   ;Initialization of click-drag range
   $continLoop = True
   $exit = False
   $x1 = 849
   $y1 = 330
   $x2 = 1325
   $y2 = 348
   While $continLoop
	  WinWaitActive("Thermo Xcalibur Qual Browser")
	  MouseClickDrag("primary", $x1, $y1, $x2, $y2) ;select the first range
	  $cmd = InputBox("Initialize AutoAvgQualBrowser", "Does this look good? Enter one of the following commands:" & @LF & "g: looks good, b: too big, s: too small, l: move range left, r: move range right, x: exit")
	  CatchErrors()
	  If $cmd == "g" Then
		 $continLoop = False
	  ElseIf $cmd == "b" Then
		 $x1 = $x1 + 20
		 $x2 = $x2 - 20
	  ElseIf $cmd == "s" Then
		 $x1 = $x1 - 20
		 $x2 = $x2 + 20
	  ElseIf $cmd == "l" Then
		 $x1 = $x1 - 20
		 $x2 = $x2 - 20
	  ElseIf $cmd == "r" Then
		 $x1 = $x1 + 20
		 $x2 = $x2 + 20
	  ElseIf $cmd == "x" Then
		 $exit = True
	  Else
		 MsgBox(0,"Initialize AutoAvgQualBrowser","Did not recognize your command. Please enter another.")
	  EndIf
   WEnd

   ;Get the starting & ending spectrum number, other variables
   If not $exit Then
	  $NumSpectToAvg = InputBox("Initialize AutoAvgQualBrowser", "Enter the number of spectra you are averaging together. Default (if left blank): 3.")
	  CatchErrors()
	  If $NumSpectToAvg == "" Then
		 $NumSpectToAvg = 3
	  EndIf
	  $ScanInterval= InputBox("Initialize AutoAvgQualBrowser", "Enter the counting interval of spectra, e.g. enter 4, if full/SIM scans occur every 4th scan."& @LF & "Default (if left blank): 1.")
	  CatchErrors()
	  If $ScanInterval == "" Then
		 $ScanInterval = 1
	  EndIf

	  $StartSpect = ""
	  ;$FinalSpect = ""
	  $StartRetT = ""
	  $FinalRetT = ""
	  While $StartSpect == ""
		 $StartSpect = InputBox("Initialize AutoAvgQualBrowser","Enter the starting spectrum number in the selected range.")
		 CatchErrors()
	  WEnd
	  While $StartRetT == ""
		 $StartRetT = InputBox("Initialize AutoAvgQualBrowser","Enter the starting retention time of the selected range.")
		 CatchErrors()
	  WEnd
	  ;While $FinalSpect == ""
		; $FinalSpect = InputBox("Initialize AutoAvgQualBrowser","Enter the final spectrum you would like to process." & @LF & "It's okay to move around in the chromatogram view. The program will reselect the range before starting.")
		 ;CatchErrors()
	  ;WEnd
	  While $FinalRetT == ""
		 $FinalRetT = InputBox("Initialize AutoAvgQualBrowser","Enter the retention time at which to stop processing." & @LF & "It's okay to move around in the chromatogram view. The program will reselect the range before starting.")
		 CatchErrors()
	  WEnd

	  ;Get the average min per spectrum - first half
	  $minPerSpectOrig = 0.13316
	  $FinalSpectNum = ""
	  WinWaitActive("Thermo Xcalibur Qual Browser")
	  ;Go to the end retention time
	  MouseClick("primary", 1905,153) ;Click the chromatogram tack button
	  MouseClick("secondary",1116,296)
	  sleep(10)
	  MouseClick("primary",1163,444)
	  WinWaitActive("Chromatogram Ranges")
	  Send(($FinalRetT/2-5) & "-" & ($FinalRetT/2+5))
	  Send("{ENTER}",0)
	  Do
		 sleep(100)
		 $color = PixelGetColor(1605,167)
	  Until $color == 15921906 ;wait until pixel in the spectrum window is grey
	  MouseClick("primary", 1904,593) ;Click the spectrum tack button
	  MouseClick("primary",1131,424)
	  $MidSpect = ""
	  While $MidSpect == ""
		 $MidSpect = InputBox("Initialize AutoAvgQualBrowser","Enter this spectrum number. This allows for calculating the average minutes per spectrum.")
		 CatchErrors()
	  WEnd
	  $MidRetT = ""
	  While $MidRetT == ""
		 $MidRetT = InputBox("Initialize AutoAvgQualBrowser","Enter the retention time for this spectrum. This allows for calculating the average minutes per spectrum.")
		 CatchErrors()
	  WEnd
	  $minPerSpectFirstHalf = ($StartRetT - $MidRetT) /($StartSpect - $MidSpect)
	  MsgBox(0,"Initialize AutoAvgQualBrowser", "The average minutes per spectrum for first half is " & $minPerSpectFirstHalf & ".")

	  ;second half
	  ;Go to the end retention time
	  MouseClick("primary", 1905,153) ;Click the chromatogram tack button
	  MouseClick("secondary",1116,296)
	  sleep(10)
	  MouseClick("primary",1163,444)
	  WinWaitActive("Chromatogram Ranges")
	  Send(($FinalRetT-5) & "-" & ($FinalRetT+5))
	  Send("{ENTER}",0)
	  Do
		 sleep(100)
		 $color = PixelGetColor(1605,167)
	  Until $color == 15921906 ;wait until pixel in the spectrum window is grey
	  MouseClick("primary", 1904,593) ;Click the spectrum tack button
	  MouseClick("primary",1131,424)
	  $FinalSpectNum = ""
	  While $FinalSpectNum == ""
		 $FinalSpectNum = InputBox("Initialize AutoAvgQualBrowser","Enter this spectrum number. This allows for calculating the average minutes per spectrum.")
		 CatchErrors()
	  WEnd
	  $FinalRetT = ""
	  While $FinalRetT == ""
		 $FinalRetT = InputBox("Initialize AutoAvgQualBrowser","Enter the retention time for this spectrum. This allows for calculating the average minutes per spectrum.")
		 CatchErrors()
	  WEnd
	  $minPerSpectSecHalf = ($MidRetT - $FinalRetT) /($MidSpect - $FinalSpectNum)
	  MsgBox(0,"Initialize AutoAvgQualBrowser", "The average minutes per spectrum for second half is " & $minPerSpectSecHalf & ".")

	  MsgBox(0,"AutoAvgQualBrowser", "You finished initializing the program. Thanks!" & @LF & "The main program is going to start now. Make sure the folder with the files is behind Qual Browser, which should be immediately behind this window." & @LF & "The results will be saved to C:\Xcalibur\Data.")
	  CatchErrors()
	  ProcessClose("QualBrowser.exe")
	  WinWaitClose("Thermo Xcalibur Qual Browser")

	  ;File-select loop:
	  For $i = 1 to $NumFiles Step 1
		 If $i <> 1 Then ;For all but the first file, hit down and open the next file
			Send("{DOWN}",0)
		 EndIf

		 ;Retention-time-select loop, 30 averages, then dump the memory and start the next portion:
		 $minPerSpect = $minPerSpectFirstHalf ;reinitialize for each file
		 $curStartSpect = $StartSpect
		 $curStartRetT = $firstStartT ;used to set range upon reopening qual browser
		 $curEndRetT = $firstEndT ;used to set range upon reopening qual browser
		 $endRangeRetT = $StartRetT + $minPerSpect*($NumSpectToAvg-1)*$ScanInterval ;used to check whether we've passed the end of the desired r.t. range, yet
		 While $endRangeRetT < $FinalRetT and not $exit
			;(re)open the file
			sleep(100)
			Send("{ENTER}",0)
			$opened = WinWaitActive("Thermo Xcalibur Qual Browser", "", 10) ;Wait for the window to open. If it doesn't call an error message.
			If Not $opened Then
			   MsgBox(0, "AutoAvgQualBrowser", "Qual Browser appears to have failed to open. Please check for an issue.")
			   CatchErrors()
			   $exit = True
			EndIf
			WinSetState("Thermo Xcalibur Qual Browser", "", @SW_MAXIMIZE)

			;Advance markers one before selecting the range
			If $curStartSpect <> $StartSpect Then
			   $curStartRetT = $curStartRetT + $minPerSpect
			   $curEndRetT = $curEndRetT + $minPerSpect
			   $endRangeRetT = $endRangeRetT + $minPerSpect
			   $curStartSpect += 1
			EndIf

			;Select the retention time range.
			MouseClick("secondary",1116,296)
			sleep(10)
			MouseClick("primary",1163,444)
			WinWaitActive("Chromatogram Ranges")
			Send($curStartRetT & "-" & $curEndRetT)
			Send("{ENTER}",0)
			Do
			   sleep(100)
			   $color = PixelGetColor(1605,167)
			Until $color == 15921906 ;wait until pixel in the spectrum window is grey
			MouseClick("primary", 1904,593) ;Click the spectrum tack button
			MouseClickDrag("primary", $x1, $y1, $x2, $y2) ;select the time retention range

			;Export-file loop
			For $j = 1 to 30 Step 1

			   ;Select the next set of spectra
			   If $j <> 1 Then
				  Send("{RIGHT}", 0)

				  ;update the minPerSpect variable
				  If $curStartSpect + $NumSpectToAvg < $MidSpect Then
					 $minPerSpect = $minPerSpectFirstHalf
				  Else
					 $minPerSpect = $minPerSpectSecHalf
				  EndIf

				  ;update markers
				  $curStartRetT = $curStartRetT + $minPerSpect
				  $curEndRetT = $curEndRetT + $minPerSpect
				  $endRangeRetT = $endRangeRetT + $minPerSpect
				  $curStartSpect += 1

				  ;wait until pixel in the spectrum window is not white
				  Do
					 sleep(100)
					 $color = PixelGetColor(1059,754)
				  Until $color <> 16777215
			   EndIf

			   If $endRangeRetT < $FinalRetT and not $exit Then
				  ;Export the averaged spectra
				  MouseClick("secondary", 1580,596) ;Right click spectrum
				  MouseClick("primary", 1638, 672) ;Click export
				  sleep(200)
				  MouseClick("primary", 1745,707) ;Click Write to Raw File

				  ;Save As window: no folder selection, save to default location. See v0.4 if you would like to ressurect this portion.
				  WinWaitActive("Save As")
				  sleep(10)
				  Send("{RIGHT}{LEFT 4}", 0) ;get back to the file name fields
				  If $NumSpectToAvg == 1 Then
					 Send("_" & $curStartSpect & "{ENTER}", 0) ;enter spectrum numbers
				  ElseIf $NumSpectToAvg == 2 Then
					 Send("_Avg" & $curStartSpect & "," & ($curStartSpect + $ScanInterval) & "{ENTER}", 0) ;enter spectrum numbers
				  Else
					 Send("_Avg" & $curStartSpect & "-" & ($curStartSpect + (($NumSpectToAvg - 1)*$ScanInterval)) & "{ENTER}", 0) ;enter spectrum numbers
				  EndIf
				  sleep(500) ;wait a bit
				  While MouseGetCursor() = 15 ;wait until hour-glass cursor is gone
					 sleep(100)
				  WEnd
				  WinWaitActive("Thermo Xcalibur Qual Browser") ;wait until the Qual Browser window is active again
			   EndIf
			Next ;End of main loop

			;Exit Qual Browser, dump the memory
			ProcessClose("QualBrowser.exe")
			WinWaitClose("Thermo Xcalibur Qual Browser")

		 WEnd ;End of retention-time-select loop
	  Next ;End of file-select loop
   EndIf ;Command x for exit or end of program
EndIf ;Qual Browser didn't open or end of program
