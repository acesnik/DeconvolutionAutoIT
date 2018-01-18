;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
; File   :  XLStoXLSX
; Author :  Anthony J. Cesnik, edited by Brian L. Frey
; History:  141010 Created the program and brought to rough functionality
;           141013 Edited, so that it takes more time when opening the file
;                  Changed first-file selection to be performed by user, so
;                  varied file menus will not impact it.
;			141014 Edited, so that doesn't need FolderPath, just uses the
;					active window and starter mouse click by user.  Also
;					made it so Excel doesn't quit, but just closes the file
;					and then minimizes Excel program so that file folder is
;					front, active window
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;Note: @LF is used for newlines in autoit: line feed

;Get file count
$FileCt = InputBox("XLStoXLSX", "Make sure Excel is full screen and has no files open, then minimize it." & @LF & "How many files do you want to process?")

;Exit if Cancel is pressed
If $FileCt = "" Then
   Exit
Else
   $FileCt = Number($FileCt)
EndIf


;Program loop
MsgBox(0,"XLStoXLSX", "Entering program loop.",1)
For $i = 1 To $FileCt

   ;Select the file to process
   If $i == 1 Then
	  MsgBox(0, "XLStoXLSX", "Please select the first XLS file in the Windows Explorer list." & @LF & "Make sure the file list is the active window behind this message before hitting okay."& @LF & "Make sure the file list is in Details format and sorted by Date Modified with oldest at the top.")
   Else
	  MouseClick("primary",951,12,1) ;activate the Windows Explorer window again
	  Send("{DOWN}", 0)
	  sleep(10)
   EndIf

   ;Open the file, and wait for Microsoft Excel to open
   ;Click on File -> Save As, wait for Save As to open
   ;BUG: It's not clicking down or opening the file correctly.
   Send("{ENTER}", 0)
   WinWaitActive("Microsoft Excel")
   sleep(1000)
   MouseClick("primary",31,42,1)
   MouseClick("primary",52,98,1)

   WinWaitActive("Save As")

   ;Select XLSX format, save
   Send("{TAB}{DOWN}", 0)
   Send("{UP 3}", 0)
   Send("{ENTER 2}", 0)

   ;close Excel file, then minimize Excel window
   MouseClick("primary",1912,38,1)
   sleep(1000)
   While MouseGetCursor() = 15
	  sleep(10)
   WEnd
   MouseClick("primary",1875,10,1)
   sleep(1000)
   While MouseGetCursor() = 15
	  sleep(10)
   WEnd
   sleep(1000)

Next

MsgBox(0,"XLStoXLSX", "Finished converting the files (hopefully!).")



