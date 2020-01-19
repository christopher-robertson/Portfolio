; SAP Login and Export

; Defining variables for dates and passwords
FormatTime, Date, , M/d/yyyy
FormatTime, Time, , h:mm tt
sap_pass := "passpass"
zexprel_pass := "password"

; Launching SAP app and entering credentials
!s::
Run, C:\Users\mypath\AppData\Local\Citrix\SelfService\Program Files\SelfService.exe
Sleep 7500
SendInput %sap_pass%{Enter}
Sleep 6000
MouseClick, Left, 560, 252
Sleep 30000
SendInput {Down 11}{Enter}
Sleep 6000
SendInput crobertson`t
Sleep 200
SendInput %zexprel_pass%
Sleep 100
SendInput {Enter}
Sleep 2000
SendInput zexprel {Enter}
Sleep 1000
MouseClick, Left, 715, 254
Sleep 1000
Return


; ZXPREL Export
; For some insane reason, this SAP application does not export data. This script iterates through the data, copying and pasting to an MS Excel sheet.

F4::
Xl := ComObjActive("Excel.Application")
variable_x := 3
variable_y := 3
variable_z := 3

Xl.Range("A1"):= "ZEXPREL"
Xl.Range("B1"):= Date
Xl.Range("C1"):= Time
Xl.Range("A2").Value := "Customer"
Xl.Range("B2").Value := "Cube"
Xl.Range("C2").Value := "Dollars"
Sleep 500

Loop 3
{
	scrolling_variable := 0
	Loop 28
	{
		SendInput `t`t`t
		Sleep 25
		SendInput {down %scrolling_variable%}
		Sleep 50
		SendInput {space}
		Sleep 4000
		Click, 394, 186
		Sleep 25
		Clipsaved := clipboardall
		Clipboard =
		Send {Home}
		Send +{End}
		Send ^c
		Clipwait
		XL.Range("A" . variable_x).Value := Clipboard
		Sleep 25
		Click, 598, 208
		Sleep 25
		Clipsaved := clipboardall
		Clipboard =
		Send {Home}
		Send +{End}
		Send ^c
		Clipwait
		XL.Range("B" . variable_y).Value := Clipboard
		Sleep 25
		Click, 911, 208
		Sleep 25
		Clipsaved := clipboardall
		Clipboard =
		Send {Home}
		Send +{End}
		Send ^c
		Clipwait
		XL.Range("C" . variable_z).Value := Clipboard
		Send {F3}
		Sleep 2000
		scrolling_variable++
		variable_x++
		variable_y++
		variable_z++
	}
	SendInput `t`t
	Sleep 50
	SendInput {PgDn}
	Sleep 2500
	SendInput +`t+`t
	Sleep 50
}
Exitapp


; CANMEX Export
F6::
Xl := ComObjActive("Excel.Application")
canmex_var_x := 3
canmex_var_y := 3
canmex_var_z := 3
canmex_var_zz := 3

Clipsaved := clipboardall
Clipboard =
Send {Home}
Send +{End}
Send ^c
Clipwait
if Clipboard in east,west
{
	Xl.Range("F1"):= "CANADA"
	Xl.Range("G1"):= Clipboard
	Xl.Range("H1"):= Date
	Xl.Range("I1"):= Time
	Xl.Range("F2").Value := "Delivery"
	Xl.Range("G2").Value := "Material"
	Xl.Range("H2").Value := "Description"
	Xl.Range("I2").Value := "Line Volume"
	Sleep 500
	While True
	{
		MouseClick, Left, 82, 273
		Loop 27
		{
			Clipsaved := clipboardall
			Clipboard =
			Send {Home}
			Send +{End}
			Send ^c
			Clipwait
			Xl.Range("F" . canmex_var_x).Value := Clipboard
			canmex_var_x++
			SendInput {Down}
		}
	
		MouseClick, Left, 273, 273
		Loop 27
		{
			Clipsaved := clipboardall
			Clipboard =
			Send {Home}
			Send +{End}
			Send ^c
			Clipwait
			Xl.Range("G" . canmex_var_y).Value := Clipboard
			canmex_var_y++
			SendInput {Down}
		}
	
		MouseClick, Left, 390, 273
		Loop 27
		{
			Clipsaved := clipboardall
			Clipboard =
			Send {Home}
			Send +{End}
			Send ^c
			Clipwait
			Xl.Range("H" . canmex_var_z).Value := Clipboard
			canmex_var_z++
			SendInput {Down}
		}
	
		MouseClick, Left, 841, 273
		Loop 27
		{
			Clipsaved := clipboardall
			Clipboard =
			Send {Home}
			Send +{End}
			Send ^c
			Clipwait
			Xl.Range("I" . canmex_var_zz).Value := Clipboard
			canmex_var_zz++
			SendInput {Down}
		}
		Sleep 200
		SendInput {PgDn}
		Sleep 5000
	}
}
If Clipboard in reg,baja,bby
{
	Xl.Range("L1"):= "MEXICO"
	Xl.Range("M1"):= Clipboard
	Xl.Range("N1"):= Date
	Xl.Range("O1"):= Time
	Xl.Range("L2").Value := "Delivery"
	Xl.Range("M2").Value := "Material"
	Xl.Range("N2").Value := "Description"
	Xl.Range("O2").Value := "Line Volume"
	Sleep 500
	While True
	{
		MouseClick, Left, 94, 230
		Loop 32
		{
			Clipsaved := clipboardall
			Clipboard =
			Send {Home}
			Send +{End}
			Send ^c
			Clipwait
			Xl.Range("L" . canmex_var_x).Value := Clipboard
			canmex_var_x++
			SendInput {Down}
		}
	
		MouseClick, Left, 250, 230
		Loop 32
		{
			Clipsaved := clipboardall
			Clipboard =
			Send {Home}
			Send +{End}
			Send ^c
			Clipwait
			Xl.Range("M" . canmex_var_y).Value := Clipboard
			canmex_var_y++
			SendInput {Down}
		}
	
		MouseClick, Left, 423, 230
		Loop 32
		{
			Clipsaved := clipboardall
			Clipboard =
			Send {Home}
			Send +{End}
			Send ^c
			Clipwait
			Xl.Range("N" . canmex_var_z).Value := Clipboard
			canmex_var_z++
			SendInput {Down}
		}
	
		MouseClick, Left, 798, 230
		Loop 32
		{
			Clipsaved := clipboardall
			Clipboard =
			Send {Home}
			Send +{End}
			Send ^c
			Clipwait
			Xl.Range("O" . canmex_var_zz).Value := Clipboard
			canmex_var_zz++
			SendInput {Down}
		}
		Sleep 200
		SendInput {PgDn}
		Sleep 5000
	}
}
else
	MsgBox Did not recognize `"%Clipboard%`".
Return


; Used for finding click coordinates.
!y::
MouseGetPos, xpos, ypos
Msgbox, X%xpos% Y%ypos%
Return


Esc::
Exitapp