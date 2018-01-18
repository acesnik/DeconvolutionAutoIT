MsgBox(0, "Initialize AutoAvgQualBrowser", "Please open that folder in Windows Explorer, select the first file, and then hit okay.")
Send("{ENTER}",0)
$opened = WinWaitActive("Thermo Xcalibur Qual Browser", "", 2) ;Wait for the window to open.
sleep(2000)
ProcessClose("QualBrowser.exe")
sleep(500)
Send("{ENTER}",0)