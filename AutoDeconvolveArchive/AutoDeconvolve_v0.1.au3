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

; 1)  Assuming the program is open in full-screen mode, click home-base, just
;     left of the y-axis of the chromatogram.
MouseClick("primary", 25 , 20, 1)

; 2)  Click right the number of initial clicks and then iteratively to the end
;     of the chromatogram, click 3, 6, 9, 12 ... 3n clicks, assuming the
;     SIM to SIM click count is 3.

; ToDo: Implement using GUI selections.

; 2a) Each time you reach a SIM scan,
;       hit the Process tab,
;       wait 5-10 sec,
;       click "Process" in the top right,
;       right click [Results] table top left corner -> "Export All,"
;       type spectrum # as additional info into title
;       save

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