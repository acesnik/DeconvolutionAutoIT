$StartRetT = 39.2
$minPerSpect = 0.13
$NumSpectToAvg = 3
$ScanInterval = 1
;$endRangeRetT = $StartRetT + minPerSpect*($NumSpectToAvg-1)*$ScanInterval ;This doesn't work! Ugh
$endRangeRetT = $StartRetT + $minPerSpect* ($NumSpectToAvg - 1) * $ScanInterval ;forgot the CA$H