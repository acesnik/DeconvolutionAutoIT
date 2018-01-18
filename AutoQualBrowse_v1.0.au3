;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
; File   :  AutoAvgQualBrowser
; Author :  Anthony J. Cesnik
; History:  141121 v1.0 implemented the solution to the crashing problem of just
;                       advancing the spectra over until the point where we left
;                       off
;
; Credit :  Brian Frey did the heavy lifting of figuring out the best workflow
;           in Protein Deconvolution/Qual Browser for current purposes!
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#include <Constants.au3>

;;;;;;INITIALIZATION;;;;;;;

Global $NumFiles = InitVar("Make a folder with only the raw files you would like to process. How many files are there?")
Global $NumSpectToAvg = InputBox("Initialize AutoAvgQualBrowser", "Enter the number of spectra you will be averaging together. Default (if left blank): 3.", 3)
Global $ScanInterval = InputBox("Initialize AutoAvgQualBrowser", "Enter the counting interval of spectra, e.g. enter 4, if full/SIM scans occur every 4th scan.", 1)

$qualBrowsing = ProcessWait("QualBrowser.exe", 1)
If $qualBrowsing Then
   MsgBox(0, "Initialize AutoAvgQualBrowser", "Close Qual Browser, please. But keep that folder in mind!")
EndIf

Global $ParamList[$NumFiles]

;PREPROCESSING
MsgBox(0, "Initialize AutoAvgQualBrowser", "Please open that folder in Windows Explorer, select the first file, and then hit okay.")
$FolderWinName = WinGetTitle("[ACTIVE]")
$opened = OpenFile()
if $opened then
   for $i = 1 to $NumFiles step 1
	  if $i <> 1 Then
		 WinWaitActive($FolderWinName)
		 Send("{DOWN}",0)
		 OpenFile()
	  EndIf

	  $startRetT = InputBox("Initialize AutoAvgQualBrowser", "Enter a starting retention time in minutes.", 35)

	  ;initialize click-drag range
	  Local $clickDragCoord[4] = [849,330,1325,348] ;x1 y1 x2 y2
	  SelectRTs($startRetT-0.5, $startRetT+0.5, False, $clickDragCoord)
	  MouseClick("primary", 1904, 593,1,1) ;Click the spectrum tack button
	  $clickDragCoord = CheckClickDrag($clickDragCoord)
	  $startSpect = InitVar("Enter the starting spectrum number in the selected range.")
	  $finalSpect = InitVar("Enter the final spectrum you would like to process.")

	  ;Store the parameters
	  Local $params[4] = [$clickDragCoord, $StartSpect, $FinalSpect, $startRetT]
	  $ParamList[$i-1] = $params

	  ;Finished initializing
	  ProcessClose("QualBrowser.exe")
   Next
EndIf

MsgBox(0,"AutoAvgQualBrowser", "You finished initializing the program. Thanks!" & @LF & "The main program is going to start now. Make sure the folder with the files is behind Qual Browser, which should be immediately behind this window." & @LF & "The results will be saved to C:\Xcalibur\Data.")

;PROCESSING
WinWaitActive($FolderWinName)
for $i = 1 to $NumFiles - 1 step 1
   Send("{UP}", 0) ;move up to the first file
next
$opened = OpenFile()
if $opened then
   for $i = 1 to $NumFiles step 1
	  if $i <> 1 Then
		 WinWaitActive($FolderWinName)
		 Send("{DOWN}",0)
		 OpenFile()
	  EndIf

	  ;variables
	  Local $params = $ParamList[$i-1]
	  Local $clickDragCoord = $params[0]
	  $startSpect = $params[1]
	  $finalSpect = $params[2]
	  $startRetT = $params[3]

	  SelectRTs($startRetT-0.5, $startRetT+0.5, True, $clickDragCoord)

	  ;process the file
	  for $curSpect = $startSpect to $finalSpect - $NumSpectToAvg + 1 step 1
		 if Mod($curSpect - $startSpect,30) == 0 and $curSpect <> $startSpect Then
			ProcessClose("QualBrowser.exe")
			WinWaitActive($FolderWinName)
			OpenFile()
			SelectRTs($startRetT-0.5, $startRetT+0.5, True, $clickDragCoord)
			AdvanceSpectra($startSpect, $curSpect)
		 EndIf
		 SaveSpect($curSpect)
		 AdvanceSpectra(0,1)
	  Next

	  ProcessClose("QualBrowser.exe")

   Next ;end of file loop
EndIf
;End of program

;;;;;;;;FUNCTION;;;;;;;;
;Name: SelectRTs
;Subroutine for selecting the retention time range with an option for performing the click drag or not.
;;;;;;;;;;;;;;;;;;;;;;;;
Func SelectRTs($StartRT, $EndRT, $performClickDrag, $clickDragCoord)
   MouseClick("primary", 1905,153,1,1) ;Click the chromatogram tack button
   MouseClick("secondary",1116,296,1,1)
   sleep(10)
   MouseClick("primary",1163,444)
   WinWaitActive("Chromatogram Ranges")
   Send($StartRT & "-" & $EndRT)
   Send("{ENTER}",0)
   Do
	  sleep(100)
	  $color = PixelGetColor(1605,167)
   Until $color == 15921906 ;wait until pixel in the spectrum window is grey
   If $performClickDrag Then
	  MouseClick("primary", 1904, 593,1,1) ;Click the spectrum tack button
	  MouseClickDrag("primary", $clickDragCoord[0], $clickDragCoord[1], $clickDragCoord[2], $clickDragCoord[3],1) ;select the time retention range
   EndIf
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: SaveSpect
;Subroutine for exporting a rawFile
;;;;;;;;;;;;;;;;;;;;;;;;
Func SaveSpect($curStartSpect)
   ;Export the averaged spectra
   MouseClick("secondary", 1580,596,1,1) ;Right click spectrum
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
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: OpenFile
;Opens the selected file in the active folder window
;;;;;;;;;;;;;;;;;;;;;;;;
Func OpenFile()
   $opened = WinWaitActive("Thermo Xcalibur Qual Browser", "", 2) ;Wait for the window to open.
   While Not $opened ;160105 Fixed bug where one or two "ENTER"s didn't successfully open QualBrowser again
	  WinActivate("[CLASS:explorer]")
	  Send("{ENTER}",0) ;Opens the first folder in qual browser
	  sleep(500)
	  $opened = WinWaitActive("Thermo Xcalibur Qual Browser", "", 2)
   WEnd
   WinSetState("Thermo Xcalibur Qual Browser", "", @SW_MAXIMIZE)
   return $opened
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: InitVar
;Initializes a variable with a while loop and a given message
;;;;;;;;;;;;;;;;;;;;;;;;
Func InitVar($msg)
   $var = ""
   While $var == ""
	  $var = InputBox("Initialize AutoAvgQualBrowser", $msg)
   WEnd
   return $var
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: AdvanceSpectra
;Moves the spectrum range right to get to the current spectrum again
;;;;;;;;;;;;;;;;;;;;;;;;
Func AdvanceSpectra($startSpect, $curSpect)
   for $i = $startSpect to $curSpect - 1 step 1
	  Send("{RIGHT}", 0)
	  Do
		 sleep(500)
		 $color = PixelGetColor(1059,754)
	  Until $color <> 16777215 ;wait until pixel in the spectrum window is not white
   next
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: CheckClickDrag
;Initializes the clickDrag region to user customization
;Takes in an array with x1,y1,x2,y2
;;;;;;;;;;;;;;;;;;;;;;;;
Func CheckClickDrag($clickDragCoord)
   $continLoop = True
   While $continLoop
	  WinWaitActive("Thermo Xcalibur Qual Browser")
	  MouseClickDrag("primary", $clickDragCoord[0], $clickDragCoord[1], $clickDragCoord[2], $clickDragCoord[3],1) ;select the first range
	  $cmd = InputBox("Initialize AutoAvgQualBrowser", "Does this look good? Enter one of the following commands:" & @LF & "g: looks good, b: too big, s: too small, l: move range left, r: move range right")
	  If $cmd == "g" Then
		 $continLoop = False
	  ElseIf $cmd == "b" Then
		 $clickDragCoord[0] += 20
		 $clickDragCoord[2] -= 20
	  ElseIf $cmd == "s" Then
		 $clickDragCoord[0] -= 20
		 $clickDragCoord[2] += 20
	  ElseIf $cmd == "l" Then
		 $clickDragCoord[0] -= 20
		 $clickDragCoord[2] -= 20
	  ElseIf $cmd == "r" Then
		 $clickDragCoord[0] += 20
		 $clickDragCoord[2] += 20
	  Else
		 MsgBox(0,"Initialize AutoAvgQualBrowser","Did not recognize your command. Please enter another.")
	  EndIf
   WEnd
   return $clickDragCoord
EndFunc
