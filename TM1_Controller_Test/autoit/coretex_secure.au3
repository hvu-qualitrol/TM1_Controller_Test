#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.6.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------



Dim $secure_window_found = False
Dim $startTime = TimerInit()

while ( Not $secure_window_found and TimerDiff($startTime) < 60000 )
	if WinExists("Cortex-M identification") Then
		ControlClick("Cortex-M identification", "", "[CLASS:Button; INSTANCE:1]")
		$secure_window_found = True
	Else
		sleep(2000)
	EndIf
WEnd	






