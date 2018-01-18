;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
; File   :  AutoDeconvolve_v0.2
; Author :  Anthony J. Cesnik
; History:  141002 v0.1 created to build framework of program.
;           141003      added framework code
;           141007      added locatoins of buttons
;           141008 v0.2 added locations of clicks and changed strategy
;           141009      edited and tested, brought to functionality (hopefully
;                  v0.3 put back into GUI
;           141010 v0.4 simplified open raw file method
;           141013      modified so that it will not skip files in list
;                       modified the save-as loop, so it waits after hopefully
;                       opening
;           141014      changed the bottom file click to fix a bug where it
;                       didn't work with more than 40 files, added more wait
;                       before clicking process
;           141015      added a pixel-color wait after processing
;                  v0.5 the program is working much more consistently with the
;                       pixel-color-based waiting. Cleaned up notes.
;           141023 v0.6 Changed the wait function after processing to include
;                       looking at the cursor - allows possibility of there
;                       being no results.
;           141122 v0.7 Made code easier to read. Replaced the GUI with MsgBoxes
;                       and InputBoxes. Added option to start at first file.
;                       Tried to make it so the program doesn't hang if a filename already
;                       exists, which is happening when the scroll bar gets slightly
;                       off somehow.
;
; Credit :  Brian Frey did the heavy lifting of figuring out the best workflow
;           in Protein Deconvolution for current purposes!
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;;;;;INITIALIZATION;;;;;;;
MsgBox(0,"AutoDeconvolve", "Open Protein Deconvolution and open your RAW data directory.")
Global $FileCt = InitVar("What is the total number of raw files in this folder?")
Global $StartFile = InitVar("Which file would you like to start at? Options: 1 or 2 for first or second file in the list.")
While $StartFile <> 1 and $StartFile <> 2
   $StartFile = InitVar("Which file would you like to start at? Options: 1 or 2 for first or second file in the list.")
WEnd
Global $StartCoord[2]
if $StartFile == 1 Then
   $StartCoord[0] = 793
   $StartCoord[1] = 324
elseif $StartFile == 2 Then
   $StartCoord[0] = 738
   $StartCoord[1] = 340
else
   MsgBox(0,"AutoDeconvolve Error", "Something's wrong with the start file initialization code.")
   Exit
EndIf

;;;;;;PROCESSING;;;;;;
;load a file and method to prime the pump
WinActivate("Thermo Protein Deconvolution")
WinWaitActive("Thermo Protein Deconvolution")
WinSetState("Thermo Protein Deconvolution", "", @SW_MAXIMIZE)
SelectCurFile(0)
LoadMethod()
BackToMethods()
MsgBox(0, "AutoDeconvolve", "Processing " & $FileCt & " files.", 1)

For $i = 1 To $FileCt - $StartFile + 1
   SelectCurFile($i)
   LoadMethod()
   Process()
   $wasNew = ExportAll()
   If $wasNew Then ;Skips the file if it hits a duplicate (be careful -- I've seen this program skip files. Hopefully I've avoided that, but it might happen again.)
	  ExportTopLevel()
   EndIf
   BackToMethods()
Next ;End of FOR program loop
MsgBox(0, "AutoDeconvolve", "Finished processing the files (hopefully!).")
;;;;;;END OF PROCESSING;;;;;;;

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
;Name: SelectCurFile
;Selects the file from the raw list
;;;;;;;;;;;;;;;;;;;;;;;;
Func SelectCurFile($i)
   MouseClick("primary",1235,338,$FileCt/30, 1) ;scroll up in case the list was not set exactly right, this will always be enough clicks because each click in the side bar goes up 41 files, independent of list length
   MouseWheel("up", 10)
   MouseClick("primary", $StartCoord[0] , $StartCoord[1], 1, 1)
   sleep(100)
   For $j = 2 To $i
	  Send("{DOWN}", 0)
	  sleep(20)
   Next
   MouseClick("primary", 1133, 923, 1, 1) ;click load
   sleep(10)

   ;get out of the check window called "Protein Deconvolution"
   $popup = WinWaitActive("Protein Deconvolution", "", 1)
   if $popup then
	  Send("{LEFT}{ENTER}",0)
	  WinWaitActive("Thermo Protein Deconvolution")
   EndIf
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: LoadMethod
;Loads the top method in the methods list
;;;;;;;;;;;;;;;;;;;;;;;;
Func LoadMethod()
   MouseClick("primary", 1337, 285, 1, 5) ;click top method
   Do
	  sleep(10)
	  $color = PixelGetColor(1403, 279)
   Until $color == 14542061
   MouseClick("primary", 1337, 285, 1, 5) ;click top method

   MouseClick("primary", 1728, 926, 1, 5) ;click load method button

   ;wait until pixel changes to white from Methods -> Process screen
   Do
	  sleep(10)
	  $color = PixelGetColor(154,760)
   Until $color == 16777215
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: Process
;Process tab will come up automatically. Click "Process" in top right, wait for screen to respond, then continue.
;;;;;;;;;;;;;;;;;;;;;;;;
Func Process()
   ;Check that window is not white and semi-frozen
   $color1 = PixelGetColor(1564,145)
   ;Wait for the results to clear before processing new spectrum
   $color2 = PixelGetColor(352,885)
   While ($color1 <> 15592941 And $color2 <> 16777215)
	  sleep(100)
	  $color1 = PixelGetColor(1564,145)
	  $color2 = PixelGetColor(352,885)
   WEnd

   ;Check that process button is active
   MouseMove(1673,149)
   sleep(200)
   $color3 = PixelGetColor(1673, 149)
   While ($color3 <> 16382457)
	  sleep(100)
	  $color3 = PixelGetColor(1673, 149)
   WEnd

   ;click process in the top right and wait for it to finish
   MouseClick("primary", 1707, 149, 1, 0)
   sleep(2000)
   $k = 0
   $color = 16777215
   While ($k < 40 And $color == 16777215) ;color is white originally
	  sleep(500)
	  $k = $k + 1
	  $color = PixelGetColor(352,885)
   WEnd
   While MouseGetCursor() = 15
	  sleep(100)
   WEnd
   sleep(500) ;extra wait for the occassional flicker
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: ExportAll
;Exports the deconvolution results as .xls file
;right click [Results] table top left corner -> "Export All," left arrow, type "all_", send {ENTER}
;;;;;;;;;;;;;;;;;;;;;;;;
Func ExportAll()
   MouseClick("secondary", 301, 860, 2) ;right click [Results]
   MouseClick("primary", 356, 886, 1) ;click "export all"
   WinWaitActive("Save As") ;Wait forever if stuck somewhere else
   Send("{LEFT 100}all_{ENTER}",0)
   Do
	  $active = WinWaitActive("Thermo Protein Deconvolution","", 5)
	  if not $active then
		 $text = WinGetText("[ACTIVE]")
		 Local $splitText = StringSplit($text, @LF)
		 If $splitText[5] == "Do you want to replace it?" Then
			Send("{RIGHT}{ENTER}",0)
			WinClose("Save As")
			return False
		 EndIf
	  EndIf
   Until $active
   return true
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: ExportTopLevel
;Exports the deconvolution results as top level .xls file
;right click [Results] table top left corner -> "Export Top Level," send {ENTER}
;;;;;;;;;;;;;;;;;;;;;;;;
Func ExportTopLevel()
   WinWaitActive("Thermo Protein Deconvolution")
   MouseClick("secondary", 301, 860, 2) ;right click [Results]
   MouseClick("primary", 356, 875, 1) ;click export top level
   WinWaitActive("Save As") ;Wait forever if stuck somewhere else
   Send("{ENTER}", 0)
   WinWaitActive("Thermo Protein Deconvolution")
EndFunc

;;;;;;;;FUNCTION;;;;;;;;
;Name: BackToMethods
;Goes back to the methods tab.
;;;;;;;;;;;;;;;;;;;;;;;;
Func BackToMethods()
   MouseClick("primary", 32, 97, 1) ;click methods tab
   Do
	  sleep(10)
	  $color = PixelGetColor(154,760)
   Until $color <> 16777215 ;wait until pixel changes to blue from Process -> Methods screen
EndFunc