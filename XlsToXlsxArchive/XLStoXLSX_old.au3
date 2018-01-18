;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;
; File   :  XLStoXLSX
; Author :  Anthony J. Cesnik
; History:  141010 Created the program and brought to rough functionality
;           141013 Edited, so that it takes more time when opening the file
;                  Changed first-file selection to be performed by user, so
;                  varied file menus will not impact it.
;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;Get folder path and file count
$FolderPath = InputBox("XLStoXLSX", "What folder path would you like to enter?")
$FileCt = InputBox("XLStoXLSX", "The address bar of a full-screen Windows Explorer window must be visible right now." & @LF & "Make sure the file list is in Details format and that it starts with an XLS file." & @LF & "How many files do you want to process?")

;Exit if Cancel is pressed
If $FolderPath = "" Or $FileCt = "" Then
   Exit
Else
   $FileCt = Number($FileCt)
EndIf

;Click the address bar, type the folder name.
;Note: @LF is used for newlines in autoit: line feed
MouseClick("primary",794,40,1)
Send($FolderPath)
Send("{ENTER}", 0)

;Program loop
MsgBox(0,"XLStoXLSX", "Entering program loop.",1)
For $i = 1 To $FileCt

   ;Select the file to process
   If $i == 1 Then
	  MsgBox(0, "XLStoXLSX", "Please select the first XLS file in the Windows Explorer list." & @LF & "Make sure the file list is the active window behind this message before hitting okay.")
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
   sleep(2000)
   MouseClick("primary",31,42,1)
   MouseClick("primary",52,98,1)

   WinWaitActive("Save As")

   ;Select XLSX format, save
   Send("{TAB}{DOWN}", 0)
   Send("{UP 3}", 0)
   Send("{ENTER 2}", 0)

   ;Exit Excel
   MouseClick("primary",1914,12,1)
   sleep(1000)
   While MouseGetCursor() = 15
	  sleep(10)
   WEnd
   sleep(1000)

Next

MsgBox(0,"XLStoXLSX", "Finished converting the files (hopefully!).")



