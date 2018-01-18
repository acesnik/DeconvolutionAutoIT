Local $StartCoord[2]
$StartCoord[0] = 793
   $StartCoord[1] = 324
   $FileCt = 2395
For $i = 1 To $FileCt
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
Next