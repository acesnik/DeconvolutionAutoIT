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
;
; Credit :  Brian Frey did the heavy lifting of figuring out the best workflow
;           in Protein Deconvolution for current purposes!
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; 0)  Initial setup: GUI to take in this info, then a Go-button
;       initial number of clicks to get from home base to a target SIM scan
;       which scans to collect data: 1, 2, 3, 4: open-ended radio boxes
;       --this is to allow for SIM and FullFT scans to be collected together
;	    take in the initial scan-number to increment and record in title

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
	  ; 2)  Process tab will come up automatically. Click "Process" in top right, wait for screen to respond, then continue.
	  ;       right click [Results] table top left corner -> "Export All,"
	  ;       left arrow, type "all_", send {ENTER}
	  ;		right click [Results] table top left corner -> "Export Top Level,"
	  ;		send {ENTER}
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
		 MouseClick("primary", 1129, 895, 1) ;click load
		 sleep(10)

		 ;get out of the check window called "Protein Deconvolution"
		 WinWaitActive("Protein Deconvolution")
		 Send("{LEFT}{ENTER}",0)
		 WinWaitActive("Thermo Protein Deconvolution")
		 MouseClick("primary", 1337, 285, 1) ;click top method
		 MouseClick("primary", 1700, 901, 1) ;click load method button

		 ;wait until pixel changes to white from Methods -> Process screen
		 Do
			sleep(10)
			$color = PixelGetColor(154,760)
		 Until $color == 16777215
		 sleep(100) ;a little extra wait

		 ;wait for possible hour-glass cursor, active window, and a bit more before clicking process
		 ;sleep(2000)
		 ;While MouseGetCursor() = 15
			;sleep(10)
		 ;WEnd
		 ;WinWaitActive("Thermo Protein Deconvolution")
		 ;sleep(1000)

		 ;click process in the top right and wait for it to finish
		 MouseClick("primary", 1707, 149, 1)
		 Do
			sleep(500) ;longer wait to take up fewer cpu cycles
			$color = PixelGetColor(352,885)
		 Until $color <> 16777215 ;wait until the pixel in the first row of the "No." column changes color
		 sleep(500) ;a little extra wait for the occassional flicker
		 ;sleep(5000)
		 ;While MouseGetCursor() = 15
		 ;	sleep(10)
		 ;WEnd

		 ;Export all
		 ;$k = 0
		 ;While $k < 3
			MouseClick("secondary", 301, 860, 2) ;right click [Results]
			MouseClick("primary", 356, 886, 1) ;click "export all"
			;$k = $k + 1
		 ;WEnd
		 WinWaitActive("Save As") ;Wait forever if stuck somewhere else
		 MouseClick("primary", 183, 469, 1)
		 sleep(10)
		 Send("all_")
		 Send("{ENTER}", 0)
		 WinWaitActive("Thermo Protein Deconvolution")

		 ;Export top level
		 WinWaitActive("Thermo Protein Deconvolution")
		 ;$k = 0
		 ;While $k < 3
			MouseClick("secondary", 301, 860, 2) ;right click [Results]
			MouseClick("primary", 356, 875, 1) ;click export top level
			;$k = $k + 1
		 ;WEnd
		 WinWaitActive("Save As") ;Wait forever if stuck somewhere else
		 Send("{ENTER}", 0)
		 WinWaitActive("Thermo Protein Deconvolution")

		 ; 3) Go back to the Methods tab:
		 MouseClick("primary", 32, 97, 1) ;click methods tab

		 ;wait until pixel changes to blue from Process -> Methods screen
		 Do
			sleep(10)
			$color = PixelGetColor(154,760)
		 Until $color <> 16777215
		 ;sleep(100)
		 ;WinWaitActive("Thermo Protein Deconvolution")

	  Next ;End of FOR program loop
	  MsgBox(0, "AutoDeconvolve", "Finished processing the files (hopefully!).")
   EndSwitch
WEnd ;End of GUI block

