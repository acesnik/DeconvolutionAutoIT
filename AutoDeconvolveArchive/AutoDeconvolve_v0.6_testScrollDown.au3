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
;
; Credit :  Brian Frey did the heavy lifting of figuring out the best workflow
;           in Protein Deconvolution for current purposes!
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; 0)  Initial setup: GUI prints information about program and takes in the
;                    number of files in folder to be processed

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=c:\users\anthony\documents\autoit projects\protein deconvolution automation\autodeconvolvegui.kxf
$Form1_1 = GUICreate("AutoDeconvolve", 544, 181, 192, 126)
$FileCtInput = GUICtrlCreateInput("40", 120, 8, 25, 21)
$Label1 = GUICtrlCreateLabel("Total number of raw files in this folder, including the full run.", 152, 8, 281, 17)
$Button1 = GUICtrlCreateButton("Go", 224, 144, 75, 25)
$Label3 = GUICtrlCreateLabel("Make sure Protein Deconvolution is the active window behind this one.", 88, 56, 339, 17)
$Label4 = GUICtrlCreateLabel("Readme:", 232, 40, 54, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$Label2 = GUICtrlCreateLabel("This program starts at the second file on the list and processes that one and all the files below it.", 48, 88, 454, 17)
$Label5 = GUICtrlCreateLabel("Before starting, if Protein Deconvolution has just been opened, load a file and method. No need to process it.", 16, 72, 516, 17)
$Label6 = GUICtrlCreateLabel("This program will hang if a filename already exists, but will continue after you hit overwrite or cancel.", 40, 120, 470, 17)
$Label7 = GUICtrlCreateLabel("This program will only work if the second file on the list is visible with this window open.", 56, 104, 410, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	$FileCt = GUICtrlRead($FileCtInput) ;Note, this is required to get the value from the GUI input field
	Switch $nMsg
		 Case $GUI_EVENT_CLOSE ;this closes the window after pressing the window exit button
			Exit
		 Case $Button1

	  ; *** Main program loop ***
	  ; 1)  First time: Starting at the methods tab in Manual Xtract mode:
	  MsgBox(0, "File Count", "File Count = " & $FileCt, 1)
	  For $i = 0 To $FileCt - 2 ;Note: To is inclusive

		 MsgBox(0, "AutoDeconvolve", "Processing File #" & $i + 1, 1)
		 ;click on the second raw file, go down to the file to process
		 If $i < 40 Then
			MouseClick("primary", 738 , 340, 1)
			MouseWheel("up",10) ;scroll up in case the list was not set exactly right
			MouseClick("primary", 738 , 340, 1)
			For $j = 1 To $i
			   Send("{DOWN}", 0)
			   sleep(10)
			Next
		 Else
			MouseClick("primary", 735 , 844, 1)
			Send("{DOWN}", 0)
			sleep(10)
		 EndIf
		 MouseClick("primary", 1133, 923, 1) ;click load
		 sleep(10)

		 ;get out of the check window called "Protein Deconvolution"
;		 WinWaitActive("Protein Deconvolution")
;		 Send("{LEFT}{ENTER}",0)
;		 WinWaitActive("Thermo Protein Deconvolution")
;		 MouseClick("primary", 1337, 285, 1) ;click top method
;		 MouseClick("primary", 1728, 926, 1) ;click load method button

		 ;wait until pixel changes to white from Methods -> Process screen
;		 Do
;			sleep(10)
;			$color = PixelGetColor(154,760)
;		 Until $color == 16777215
		 ;sleep(1000) ;a little extra wait

		 ;Check that window is not white and semi-frozen
;		 $color1 = PixelGetColor(1564,145)
		 ;Wait for the results to clear before processing new spectrum
;		 $color2 = PixelGetColor(352,885)
;		 While ($color1 <> 15592941 And $color2 <> 16777215)
;			sleep(100)
;			$color1 = PixelGetColor(1564,145)
;			$color2 = PixelGetColor(352,885)
;		 WEnd

		 ;Check that process button is active
;		 MouseMove(1673,149)
;		 sleep(200)
;		 $color3 = PixelGetColor(1673, 149)
;		 While ($color3 <> 16382457)
;			sleep(100)
;			$color3 = PixelGetColor(1673, 149)
;		 WEnd

		 ; 2)  Process tab will come up automatically. Click "Process" in top right, wait for screen to respond, then continue.
		 ; right click [Results] table top left corner -> "Export All," left arrow, type "all_", send {ENTER}
		 ; right click [Results] table top left corner -> "Export Top Level," send {ENTER}

		 ;click process in the top right and wait for it to finish
;		 MouseClick("primary", 1707, 149, 1)
;		 sleep(2000)
;		 $k = 0
;		 While ($k < 40 And $color == 16777215) ;color is white originally
;			sleep(500)
;			$k = $k + 1
;			$color = PixelGetColor(352,885)
;		 WEnd
;		 While MouseGetCursor() = 15
;			sleep(100)
;		 WEnd
;		 sleep(500) ;extra wait for the occassional flicker

		 ;Export all
;		 MouseClick("secondary", 301, 860, 2) ;right click [Results]
;		 MouseClick("primary", 356, 886, 1) ;click "export all"
;		 WinWaitActive("Save As") ;Wait forever if stuck somewhere else
;		 MouseClick("primary", 132, 455, 1)
;		 sleep(10)
;		 Send("all_")
;		 Send("{ENTER}", 0)
;		 WinWaitActive("Thermo Protein Deconvolution")

		 ;Export top level
;		 WinWaitActive("Thermo Protein Deconvolution")
;		 MouseClick("secondary", 301, 860, 2) ;right click [Results]
;		 MouseClick("primary", 356, 875, 1) ;click export top level
;		 WinWaitActive("Save As") ;Wait forever if stuck somewhere else
;		 Send("{ENTER}", 0)
;		 WinWaitActive("Thermo Protein Deconvolution")

		 ; 3) Go back to the Methods tab:
;		 MouseClick("primary", 32, 97, 1) ;click methods tab
;		 Do
;			sleep(10)
;			$color = PixelGetColor(154,760)
;		 Until $color <> 16777215 ;wait until pixel changes to blue from Process -> Methods screen

	  Next ;End of FOR program loop
	  MsgBox(0, "AutoDeconvolve", "Finished processing the files (hopefully!).")
   EndSwitch
WEnd ;End of GUI block

