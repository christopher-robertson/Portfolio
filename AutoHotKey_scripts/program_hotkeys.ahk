; Run Program Hotkeys
; A shortcut of this script sits in my Windows startup folder so that these commands are available to me at all times.

; Store passwords as variables to automatically login into different accounts and applications
rp_pass := "password1"
em_pass := "password2"

; Launch other AHK scripts.
; Most of my AHK scripts have the ESC key set as a hotkey to kill the script when I'm done using it.
; 

^+e::
	Run, E:\scripts\ahk\email.ahk
	Return

^+i::
	Run, E:\scripts\ahk\inbox_scroll.ahk
	Return

^+z::
	Run, E:\scripts\ahk\sap_export.ahk
	Return

^+j::
	Run, E:\WorkLog.txt
	Return

<#c::
	Run, C:\Windows\syswow64\Windowspowershell\v1.0\powershell.exe
	Return

^+k::
	Run, E:\KeePass-2.42.1\KeePass.exe
	Return

^+p::
    Run, E:\scripts\ahk\pc_query.ahk
	Return

^+r::
	Run, C:\Program Files\RedPrairie\Client\core\DlxLaunch.exe
	Sleep 3000
	WinActivate, RedPrairie Solutions
	SendInput {Down}{Enter}
	Sleep 5000
	SendInput crobertson`t%RP_Pass%{Enter}
	Return

#m::
	Run, C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
	Sleep 2500
	SendInput https://app.easymetrics.com{Enter}
	Sleep 2500
	SendInput `t
	Sleep 25
	SendInput crobertson`t
	Sleep 25
	SendInput %em_pass%
	Sleep 50
	SendInput {Enter}
	Return

    
; Time Stamp Hotkeys

!<+t::
FormatTime, Time, , yyyy-MM-dd HH:mm
SendInput, %Time%
Return

!<+d::
FormatTime, Time, , yyyy-MM-dd
SendInput, %Time%
Return

!<+m::
FormatTime, Time, , yyMMdd
SendInput, %Time%
Return

::esig::
(
Christopher Robertson
Data Analyst
Fedex Supply Chain
909-123-4567
Fender North America Distribution
)
Return

::----::
(
-----------------------------------------------------------------------------
)
Return
