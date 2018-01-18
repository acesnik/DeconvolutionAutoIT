;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
; File   :  AutoDeconvolve_v0.1
; History:  141002 v0.1 created to build framework of program.
;           141003      added framework code
;           141007      added locatoins of buttons
;
;
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; 0)  Initial setup: GUI to take in this info, then a Go-button
;       initial number of clicks to get from home base to a target SIM scan
;       which scans to collect data: 1, 2, 3, 4: open-ended radio boxes
;       --this is to allow for SIM and FullFT scans to be collected together
;	    take in the initial scan-number to increment and record in title

#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=C:\Users\Anthony\Documents\AutoIt Projects\Protein Deconvolution Automation\AutoDeconvolveGUI.kxf
$Form1 = GUICreate("AutoDeconvolve", 486, 274, 192, 126)
$Checkbox1 = GUICtrlCreateCheckbox("Scan 1", 48, 128, 97, 17)
$Checkbox2 = GUICtrlCreateCheckbox("Scan 2", 48, 144, 97, 17)
$Checkbox3 = GUICtrlCreateCheckbox("Scan 3", 48, 160, 97, 17)
$Checkbox4 = GUICtrlCreateCheckbox("Scan 4", 48, 176, 97, 17)
$Radio1 = GUICtrlCreateRadio("Analyze every 3 scans", 16, 88, 353, 17)
$Radio2 = GUICtrlCreateRadio("Analyze below of every 4 scans following the first scan analyzed", 16, 112, 329, 17)
Local $idSpectNum = GUICtrlCreateInput("200", 16, 8, 121, 21)
$Label1 = GUICtrlCreateLabel("First scan # on current chromatogram (example: 204)", 152, 8, 252, 17)
$Button1 = GUICtrlCreateButton("Go", 208, 240, 75, 25)
$Label3 = GUICtrlCreateLabel("Reminders: make sure Protein Deconvolution is in full-screen mode", 88, 208, 318, 17)
Local $idInitClicks = GUICtrlCreateInput("0", 16, 48, 121, 21)
$Label5 = GUICtrlCreateLabel("Number of right clicks needed to advance to first scan to analyze", 152, 48, 310, 17)
$Label2 = GUICtrlCreateLabel("and the active window behind this one.", 152, 224, 189, 17)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

	EndSwitch
WEnd

; 1)  First time: Starting at the methods tab in Manual Xtract mode, click on the second raw file, click load,
;     click the first method, click load method
MouseClick("primary", 25 , 20, 1)

; *** Main program loop ***
; 2)  Process tab will come up automatically. Click "Process" in top right, wait for screen to respond, then continue.
;       right click [Results] table top left corner -> "Export All,"
;       right arrow, left arrow X4, type "_all", send {ENTER}
;		right click [Results] table top left corner -> "Export Top Level,"
;		right arrow, left arrow X4, type "_topLevel", send {ENTER}

MouseClick("primary", 20, 20, 1)
Sleep(10000)
MouseClick("primary", 20, 20, 1)
MouseClick("secondary", 20, 20, 1)
MouseClick("primary", 20, 20, 1)
Send("{RIGHT}", 0)
$i = 4 ;the number of characters to erase before the spectrum number is included
While $i > 0
   Send ("{BS}", 0)
   $i = $i - 1
WEnd
Send(GUICtrlRead($idSpectNum), 1)
MouseClick("primary", 20, 20, 1)

; 3) Go back to the Methods tab:
;      Click Methods tab, click the first file, hit down arrow an incrementing number of times to a maximum
;      Hit {LEFT}, {ENTER} to get out of the box
;      Click top method at (1337, 285) and then hit the load method button at (1700, 901)
;      Hit {LEFT}, {ENTER} to get out of the box
;      Wait for program to respond to window change?

; Externally in R: Add all of these Excel spreadsheets into a master sheet
;
; Button locations relative to active window:
; home base: 25, 436
; process tab: 624, 106
; "Process" in top right: 1707, 149
; Results table top left corner: 299, 866
; "Export all": 340, 869
; "Export top level": 347, 877
; Methods tab: 106, 109
; Location of files to select in methods tab:
; 46, -140
; - , -128
; - , -114
; - , -101
; - , -85