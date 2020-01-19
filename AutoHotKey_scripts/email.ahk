; Email Text Expansion
; I use this script to write emails for me in Outlook. I launch it from hotkeys in the program_hotkeys.ahk script

^1:: ; Attendance
SendInput ^n
Sleep 10
ysterday = %a_now%
ysterday += -1, days
FormatTime, ysterday, %ysterday%, yyyy-M-d
SendInput example.email1@foo.com`texample.email2@foo.com`texample.email3@foo.com`texample.email4@foo.com`texample.email5@foo.com`texample.email6@foo.com`texample.email7@foo.com`texample.email8@foo.com`texample.email9@foo.com`t`t`tAttendance for %ysterday%`t`r`r`r{CTRL DOWN}b{CTRL UP}Please Review the Following Teammates:`r`r{CTRL DOWN}b{CTRL UP}
Return

^2:: ; Ultras Status
SendInput ^n
Sleep 10
today = %a_now%
FormatTime, today, %ysterday%, yyyy-M-d HH:00
SendInput example.email1@foo.com`texample.email2@foo.com`texample.email3@foo.com`texample.email4@foo.com`texample.email5@foo.com`texample.email6@foo.com`texample.email7@foo.com`texample.email8@foo.com`texample.email9@foo.com`t`tUltras Status as of %today%`tHi Team,`r`rPlease see below for an update on Ultras:`r`r`r{CTRL DOWN}b{CTRL UP}Total Temp Room Status by Location:`r`r`r`rUnit Distribution for Ultra SKU{asc 0039}s:`r`r`r`rUltras Remaining outside of Temp Room: 100{asc 0037} of units in A and B locations and 0{asc 0037} still being moved from higher locations:{CTRL DOWN}b{CTRL UP}`r`r`r`rThanks,`rChris`r`r`rChristopher Robertson`rData Analyst`rSupply Chain`r909-123-4567`rNorth America Distribution
Return

^3:: ; Supervisor Report
SendInput ^n
Sleep 10
ysterday = %a_now%
ysterday += -1, days
FormatTime, ysterday, %ysterday%, yyyy-M-d
SendInput example.email1@foo.com`texample.email2@foo.com`texample.email3@foo.com`texample.email4@foo.com`texample.email5@foo.com`texample.email6@foo.com`texample.email7@foo.com`texample.email8@foo.com`texample.email9@foo.com`texample.email10@foo.com`t`t`tShift Breakdown and OT for %ysterday%`t{CTRL DOWN}b{CTRL UP}Yesterday - OT: ,    Direct Labor CPU: `r`rTrending WTD - OT: ,    Direct Labor CPU: `r`rTrending MTD - OT: ,    Direct Labor CPU: {CTRL DOWN}b{CTRL UP}`r`r`rPlease see below, teammates who worked more than 30 minutes of overtime yesterday:`r`r
Return

^4:: ; Gap Time Report
SendInput ^n
Sleep 10
ysterday = %a_now%
ysterday += -1, days
FormatTime, ysterday, %ysterday%, yyyy-M-d
SendInput example.email1@foo.com`texample.email2@foo.com`texample.email3@foo.com`texample.email4@foo.com`texample.email5@foo.com`texample.email6@foo.com`texample.email7@foo.com`t`t`t`Gap Time for %ysterday%`tPlease see attached.`r`rThanks,`rChris
Return

Esc::
Exitapp
