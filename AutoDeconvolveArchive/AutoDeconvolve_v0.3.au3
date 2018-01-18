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
$FileCt = GUICtrlCreateInput("40", 120, 8, 25, 21)
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
	Switch $nMsg
		 Case $GUI_EVENT_CLOSE ;this closes the window after pressing the window exit button
			Exit
		 Case $Button1
			; *** Main program loop ***
	  ; 2)  Process tab will come up automatically. Click "Process" in top right, wait for screen to respond, then continue.
	  ;       right click [Results] table top left corner -> "Export All,"
	  ;       left arrow, type "all_", send {ENTER}
	  ;		right click [Results] table top left corner -> "Export Top Level,"
	  ;		send {ENTER}

	  For $i = $FileCt - 1 To 1 Step -1 ;Note: To is inclusive

		 ;Open the raw file
		 If $i = $FileCt - 1 Then
			; 1)  First time: Starting at the methods tab in Manual Xtract mode:
			MouseClick("primary", 738 , 340, 1) ;click on the second raw file
			WinWaitActive("Thermo Protein Deconvolution")
			MouseClick("primary", 1129, 895, 1) ;click load
			sleep(10)

			;get out of the check window
			WinWait("Protein Deconvolution") ;check on title
			sleep(10)
			Send("{LEFT}{ENTER}",0)
			WinWaitActive("Thermo Protein Deconvolution")
			MouseClick("primary", 1337, 285, 1) ;click top method
			MouseClick("primary", 1700, 901) ;click load method button
			WinWait("Thermo Protein Deconvolution")

		 Else
			MouseClick("primary", 738 , 340, 1) ;click on the second raw file
			For $j = $FileCt - 2 To $i Step -1
			   Send("{DOWN}", 0)
			Next

			MouseClick("primary", 1125, 907, 1) ;click load, check on x,y
			WinWait("Protein Deconvolution") ;check on title
			Send("{LEFT}{ENTER}",0)
			WinWait("Thermo Protein Deconvolution")
			MouseClick("primary", 1337, 285, 1) ;click top method
			MouseClick("primary", 1700, 901) ;click load method button
			WinWait("Thermo Protein Deconvolution")
		 EndIf

		 MouseClick("primary", 1707, 149, 1) ;click process in the top right
		 sleep(5000)
		 While MouseGetCursor() = 15
			sleep(10)
		 WEnd

		 $k = 0
		 While WinWait("Save As", "", 1) == 0 Or $k < 3
			MouseClick("secondary", 301, 860, 1) ;right click [Results]
			MouseClick("primary", 356, 886, 1) ;click "export all"
			$k = $k + 1
		 WEnd

		 While MouseGetCursor() = 15
			sleep(10)
		 WEnd

		 MouseClick("primary", 132, 455, 1)
		 Send("all_")
		 Send("{ENTER}", 0)

		 WinWaitActive("Thermo Protein Deconvolution")
		 $k = 0
		 While WinWait("Save As", "", 1) == 0 Or $k < 3
			MouseClick("secondary", 301, 860, 1) ;right click [Results]
			MouseClick("primary", 356, 875, 1) ;click export top level
			$k = $k + 1
		 WEnd

		 Send("{ENTER}", 0)

		 WinWaitActive("Thermo Protein Deconvolution")

	  ; 3) Go back to the Methods tab:
	  ;      Click Methods tab, click the first file, hit down arrow an incrementing number of times to a maximum, hit load
	  ;      Hit {LEFT}, {ENTER} to get out of the box
	  ;      Click top method at (1337, 285) and then hit the load method button at (1700, 901)
	  ;      Hit {LEFT}, {ENTER} to get out of the box
	  ;      Wait for program to respond to window change?
		 MouseClick("primary", 32, 97, 1) ;click methods tab
		 While MouseGetCursor() = 15
			sleep(10)
		 WEnd

	  Next

	EndSwitch
WEnd

