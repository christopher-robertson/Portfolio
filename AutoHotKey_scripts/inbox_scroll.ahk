; Outlook Inbox email scroll
; Select the last message in Inbox and launch this with ALT-u
; Simple, but extremely effective at getting through emails. Spends two seconds per email, then moves up.

!u::
While True
{
    SendInput {Up}
    Sleep 2000
}

Esc::
Exitapp