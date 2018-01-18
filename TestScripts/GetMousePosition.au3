; source: https://www.autoitscript.com/wiki/Snippets_(_AutoIt_Mouse_&_Keyboard_)
; Author: shmuelw1
;
; This script returns the current mouse position
; Edit the value of $mode as required
#include <Misc.au3> ; required for _IsPressed

Global $mode = 0 ; this sets the MouseCoordMode - see Opt("MouseCoordMode" later in the script
Global $modeText
Global $pos

Opt("MouseCoordMode", $mode) ; 1 = absolute screen position, 0 = relative to active windows, 2 = relative to client area
Opt("TrayIconDebug", 1)

Switch $mode
    Case 0
        $modeText = "relative to the active window is:" & @CRLF
    Case 1
        $modeText = "from the top-right of the screen is:" & @CRLF
    Case 2
        $modeText = "relative to the client area of the active window is:" & @CRLF
EndSwitch

ToolTip("Move the pointer to the desired location. Press Enter to continue.", @DesktopWidth/2-250, @DesktopHeight/2-10)

While Not _IsPressed("0D") ; wait for Enter key
    Sleep(100)
WEnd

ToolTip('') ; close ToolTip

$pos = MouseGetPos()

MsgBox(0, "Mouse Position", "The mouse position " & $modeText & $pos[0] & "," & $pos[1] & " (x,y)")