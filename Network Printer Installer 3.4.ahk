#NoTrayIcon
#SingleInstance Force

If (A_IsAdmin = 0) {
	MsgBox Please reopen this script with administrative priviledges
	ExitApp
}

;********************************* schoolOne Printers and IPs **************************************************
global schoolOne := {"Library (118) - Sharp 623N"		:		"10.45.79.8"
	,		"Rm 215 - Ricoh 6002"							:		"10.45.76.2"
	,		"Main Office - HP 9050"							:		"10.45.79.2"
	,		"Main Office - Sharp 4100"						:		"10.45.79.4"
	,		"Rm 244 - Sharp MX503N"							:		"10.45.76.8"
	,		"Rm 244 - HP LaserJet 4200n"					:		"10.45.76.4"}

;********************************* schoolTwo Printers and IPs **************************************************
global schoolTwo := {"Main Office - Sharp 4140"	: 		"10.48.191.185"
	,		"Community Center - Sharp 4110"					:		"10.48.191.2"
	,		"Rm 214 (Lounge) - Sharp 623"					:		"10.48.190.204"
	,		"Rm 314 - Sharp 623"							:		"10.48.188.33"}

;********************************* schoolThree Printers and IPs **************************************************
global schoolThree := {"Rm 214 -  HP 9050"				: 		"10.4.157.3"
	,					"Rm 314 -  HP 9050"					:		"10.48.191.2"}


;********************************* LIST_OF_SCHOOLS **************************************************
schoolPrinters := {	"schoolOne"			:schoolOne
	, 				"schoolTwo"			:schoolTwo
	, 				"schoolThree"		:schoolThree}

;********************************* GUI DDL LIST**************************************************
for i in schoolPrinters {  ; generates string of school names with "|" in between, (i.e. "|schoolOne|schoolTwo|schoolThree") --- changed in 3.1
	school_DDL_list := school_DDL_list "|" . i
}

;*********************************Printer DRIVER NAMES****************************
global RicohUDName := "PCL6 Driver for Universal Print"
global SharpUDName := "SHARP UD2 PCL6"
global HpUDName := "HP Universal Printing PCL 6"
global XeroxUDName := "Xerox Global Print Driver PCL6"
global LexmarkUDName := "Lexmark Universal v2 XL"
global KyoceraUDName := "Kyocera Classic Universaldriver PCL6"
global SamsungUDName := "Samsung Universal Print Driver 3" ; or "Samsung Universal Print V3.00.10.00" (TESTING ON ACTUAL SAMSUNG MFP needed) as of 3/4/17

;********************************* DRIVER PATHS****************************
global LexmarkPath := """" A_ScriptDir "\Printer Universal Drivers\Lexmark\x86-x64\LMUD1p40.inf" """"
global KyoceraPath := """" A_ScriptDir "\Printer Universal Drivers\Ricoh\x86-x64\OEMsetup.inf" """"
global SamsungPath := """" A_ScriptDir "\Printer Universal Drivers\Ricoh\x86-x64\us00a.inf" """"

if (A_Is64bitOS = 1) {
global RicohPath := """" A_ScriptDir "\Printer Universal Drivers\Ricoh\64bit\oemsetup.inf" """"
global SharpPath := """" A_ScriptDir "\Printer Universal Drivers\Sharp\64bit\sfweMENU.inf" """"
global HpPath := 	"""" A_ScriptDir "\Printer Universal Drivers\HP\64bit\hpcu190u.inf" """"
global XeroxPath := """" A_ScriptDir "\Printer Universal Drivers\Xerox\64bit\x2UNIVX.inf" """"
}
else {
global RicohPath := """" A_ScriptDir "\Printer Universal Drivers\Ricoh\32bit\oemsetup.inf" """"
global SharpPath := """" A_ScriptDir "\Printer Universal Drivers\Sharp\32bit\sfweJENU.inf" """"
global HpPath := 	"""" A_ScriptDir "\Printer Universal Drivers\HP\32bit\hpcu190c.inf" """"
global XeroxPath := """" A_ScriptDir "\Printer Universal Drivers\Xerox\32bit\x2UNIVX.inf" """"
}

;**********************Empty Vars**************************************************************************

global ipADDRESS :=
global printerName :=
global driverName :=
global driverPath :=

;******************INSTALL PRINTER FUNCTION***********************
installPrinter() {
; details found @ http://woshub.com/manage-printers-and-drivers-from-the-command-line-in-windows-8/
;Printer driver name must be same as listed in .inf file

;install driver -####### Printer driver name must be same as listed in .inf file#######
GuiControl,Disable,InstallButton
	Progress,CWFFFFFF b w500 Fm12,, Installing: %printerName% `n @ IP Address: %ipADDRESS%
	RunWait, %comspec% /c cscript "%A_WinDir%\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs" -a -m "%driverName%" -i %driverPath%,,Hide
	Progress, 33
	Runwait, % comspec " /c  cscript " """" A_WinDir "\System32\Printing_Admin_Scripts\en-US\Prnport.vbs" """" " -a -r " ipADDRESS " -h " ipADDRESS " -o raw -n 9100",,Hide
	Progress, 66
	RunWait, %comspec% /c cscript "%A_WinDir%\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs" -a -p "%printerName%" -m "%driverName%" -r "%ipADDRESS%",,Hide
	Progress, 100
	Sleep, 500
	Progress, 100,, Installation complete!!!
	Sleep, 1500
	Progress, OFF
GuiControl,Enable,InstallButton
}

;********************Check to print test page and printer preferences****************************
;*******************GUI******************************************

;PRESET side
Menu, tray, Icon , %A_WorkingDir%\Printer Universal Drivers\mfp.ico, 1,
Gui, Font, S16 CDefault, Arial
Gui, Add, GroupBox, x12 y20 w280 h310 , Preset
Gui, Font, S14 CDefault, Arial
Gui, Add, Text, vPresetprinTitle x102 y200 w100 h30 +Center, Model
Gui, Add, DropDownList, r8 vCtitle GschoolChoice x72 y120 w160 h30 , %school_DDL_list%
Gui, Add, Text, vSchooltitle x102 y70 w100 h30 +Center, School
Gui, Add, DropDownList, r8 VprinterList x22 y250 w260 h30 ,

;CUSTOM SIDE
Gui, Font, S16 CDefault, Arial
Gui, Add, GroupBox, x302 y20 w290 h320 , Custom
Gui, Font, S14 CDefault, Arial
Gui, Add, Text, vNametitle x392 y70 w110 h20 +Center, Brand
Gui, Add, DropDownList, r7 vCusBrand x342 y100 w210 h30 ,  HP|KYOCERA|LEXMARK|RICOH/GESTETNER|SAMSUNG|SHARP|XEROX
Gui, Add, Text, vPrinTitle x387 y160 w120 h20 +Center, Name
Gui, Add, Edit, vCusName x312 y190 w270 h30,
Gui, Add, Text, vIPtitle x400 y250 w95 h30 , IP Address
Gui, Font, S16 CDefault, Arial
Gui, Add, Edit, vCusIP x352 y290 w190 h30 , 10.

;OTHER CONTROLS
Gui, Font, S16 CDefault, Arial
Gui, Add, Radio, gRadioPresetInstr vRadPreset x92 y20 w20 h30 checked,
Gui, Add, Radio, gRadioCustomInstr vRadCustom x392 y20 w20 h30,
Gui, Font, S22 CDefault, Arial
Gui, Add, Button, gSubmitAndRun vInstallButton x184 y360 w240 h50 , Install Printer
; Generated using SmartGUI Creator for SciTE
	GuiControl,Enable,PresetprinTitle
	GuiControl,Enable,Ctitle
	GuiControl,Enable,Schooltitle
	GuiControl,Enable,printerList
	GuiControl,Disable,Nametitle
	GuiControl,Disable,CusBrand
	GuiControl,Disable,PrinTitle
	GuiControl,Disable,CusName
	GuiControl,Disable,IPtitle
	GuiControl,Disable,CusIP

	GuiControl,,CusBrand,|HP|KYOCERA|LEXMARK|RICOH/GESTETNER|SAMSUNG|SHARP|XEROX
	GuiControl,,CusName,
	GuiControl,,CusIP,

;************************MENU BAR*******************MENU BAR *******************MENU BAR********************



RunWait, %comspec% /c cscript "%A_WinDir%\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs" -l | clip,,hide
Sleep 100
listprn := Clipboard

linarr := StrSplit(listprn, "`r" "`n")

PrnNamearray := []
for i in linarr {
	emptV := linarr[i]
	if emptV contains Printer name
	{
	PrnNamearray.Push(SubStr(linarr[i],14,StrLen(linarr[i])))
	}
}

for printerIndex in PrnNamearray {
    submenu := PrnNamearray[printerIndex]
    Menu, % submenu, Add, Printer Properties, SubSubMenAction
    Menu, % submenu, Add, Printer Preferences, SubSubMenAction
    Menu, % submenu, Add, Print Test Page, SubSubMenAction
    Menu, % submenu, Add
    Menu, % submenu, Add, Delete Printer, SubSubMenAction

    Menu, OptionsMenu, Add, % PrnNamearray[printerIndex], :%submenu% ; generates list of printers in "Options" menu. points to sub menu items with the name of specific printer in loop
}

Clipboard :=


Menu, FileMenu, Add,% "Devices and Printers", DevPtrs
Menu, FileMenu, Add,% "Print Management", PrntMngr
Menu, FileMenu, Add ; Divider
Menu, FileMenu, Add, Reset Printer List, ResPrnLst
Menu, FileMenu, Add, E&xit, GuiClose


Menu, OptionsMenu, Add
Menu, OptionsMenu, Add, % "Delete All Standard Windows Printers", DelDefPrtrs
Menu, OptionsMenu, Add, % "Delete All Printers", DelAllPrtrs




; Attach the sub-menus that were created above.
Menu, MyMenuBar, Add, &File, :FileMenu
Menu, MyMenuBar, Add, &Options, :OptionsMenu

Gui, Menu, MyMenuBar ; Attach MyMenuBar to the GUI

Gui, Show, w613 h441, %A_ScriptName%
return

SubSubMenAction:
for i in PrnNamearray{
	if (A_ThisMenuItem = "Printer Properties") {
		RunWait, % "rundll32.exe printui.dll`,PrintUIEntry /p /n" """" A_ThisMenu """"
		return
	}
	if (A_ThisMenuItem = "Printer Preferences") {
		RunWait % "rundll32.exe printui.dll`,PrintUIEntry /e /n " """" A_ThisMenu """"
		return
	}
	if (A_ThisMenuItem = "Print Test Page") {
		RunWait % "rundll32.exe printui.dll`,PrintUIEntry /k /n " """" A_ThisMenu """"
		return
	}
	if (A_ThisMenuItem = "Delete Printer") {
		MsgBox, 289,, % "Are you sure you want to delete " A_ThisMenu "?"
		IfMsgBox OK
			{
			RunWait % "rundll32.exe printui.dll`,PrintUIEntry /dl /n " """" A_ThisMenu """"
			MsgBox % A_ThisMenu " has been successfully deleted!!!"
			}
		return
	}
}
return
ResPrnLst:
for i, submenu in PrnNamearray {
        if (submenu = A_ThisMenu) {
            menu, % submenu, DeleteAll ; optional
            menu, OptionsMenu, Delete, % submenu
            PrnNamearray.delete(i) ; optional
            break
        }
    }
    return
DevPtrs: ;Run devices and printers
Run, control printers
return

PrntMngr: ;Run Print Management
Run mmc.exe printmanagement.msc
return

DelDefPrtrs: ; command to delete default printers installed on Windows installation
RunWait, %comspec% /c cscript "%A_WinDir%\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs" -d -p Fax,,Hide
RunWait, %comspec% /c cscript "%A_WinDir%\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs" -d -p "Microsoft XPS Document Writer",,Hide
RunWait, %comspec% /c cscript "%A_WinDir%\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs" -d -p "Send To OneNote 2013",,Hide
MsgBox, All Printers Successfully Deleted!!!
return

DelAllPrtrs: ;Command to delete all installs printers
MsgBox, 289,, % "Are you sure you want to delete all currently installed printers?"
	IfMsgBox OK
	{
    RunWait, %comspec% /c cscript "%A_WinDir%\System32\Printing_Admin_Scripts\en-US\prnmngr.vbs" -x,,Hide
	MsgBox, All Standard Windows Printers Successfully Deleted!!!
	}
return


RadioPresetInstr:
Gui, Submit,NoHide
if RadPreset {
	GuiControl,Enable,PresetprinTitle
	GuiControl,Enable,Ctitle
	GuiControl,Enable,Schooltitle
	GuiControl,Enable,printerList
	GuiControl,Disable,Nametitle
	GuiControl,Disable,CusBrand
	GuiControl,Disable,PrinTitle
	GuiControl,Disable,CusName
	GuiControl,Disable,IPtitle
	GuiControl,Disable,CusIP

	GuiControl,,CusBrand,|HP|KYOCERA|LEXMARK|RICOH/GESTETNER|SAMSUNG|SHARP|XEROX
	GuiControl,,CusName,
	GuiControl,,CusIP,
	return
}

RadioCustomInstr:
Gui, Submit,NoHide
if RadCustom { ;if "Preset" radio is checked
	GuiControl,Disable,PresetprinTitle
	GuiControl,Disable,Ctitle
	GuiControl,Disable,Schooltitle
	GuiControl,Disable,printerList
	GuiControl,Enable,Nametitle
	GuiControl,Enable,CusBrand
	GuiControl,Enable,PrinTitle
	GuiControl,Enable,CusName
	GuiControl,Enable,IPtitle
	GuiControl,Enable,CusIP

	GuiControl,,Ctitle, %school_DDL_list%
	GuiControl,,printerList, |
	return
}
else
;	goto RadioPresetInstr


schoolChoice:
Gui, Submit, NoHide

/* cycles through printer list array
changed in 3.1 -------		if (Ctitle = "schoolName") {
							GuiControl,,printerList, %printerlistDDL string%
							return
}
*/
global ddlList :=

for i in schoolPrinters{ ;**********Cycle through schoolPrinters array to create dynamic dropdown list in printers field ************* changed in 3.1
		if (Ctitle = i) {
			for k in schoolPrinters[i]{
				ddlList := ddlList "|" . k
		}
		GuiControl,,printerList, %ddlList%
		return
	}
}

SubmitAndRun:
Gui, Submit,NoHide
;**********PRESET SIDE*******************************************************************************

If RadPreset {
	if (Ctitle = "" || Ctitle = ) {
	MsgBox,16,%A_ScriptName%, % "Please select a school."
	return
}

if (printerList = "" || printerList = ) {
		MsgBox,16,%A_ScriptName%, % "Please select a printer."
		return
}
else {
;**********Cycle through schoolPrinters array to check for brand name in printer title and run respective install function ************* changed in 3.1
	for i in schoolPrinters
		for k, v  in schoolPrinters[i] {
				drv :=
			if (printerList = k ) {
				printerName := k , ipADDRESS := v
				If k contains sharp
					{
					driverName := SharpUDName , driverPath:= SharpPath
					installPrinter()
					}

				If k contains HP
					{
					driverName := HpUDName , driverPath:= HpPath
					installPrinter()
					}

				If k contains ricoh
					{
					driverName := RicohUDName , driverPath:= RicohPath
					installPrinter()
					}

				If k contains lexmark
					{
					driverName := LexmarkUDName , driverPath:= LexmarkPath
					installPrinter()
					}

				If k contains xerox
					{
					driverName := XeroxUDName , driverPath:= XeroxPath
					installPrinter()
					}

				If k contains kyocera
					{
					driverName := KyoceraUDName , driverPath:= KyoceraPath
					installPrinter()
					}

				If k contains samsung
					{
					driverName := SamsungUDName , driverPath:= SamsungPath
					installPrinter()
					}
			}
		}
	}
}

;**********Custom SIDE*************************************************************************************************************************
If RadCustom { ;if "Custom" radio is checked

errmsg := ""

ValidIP(IPAddress) ; checks if IP address is Valid
{
	if (StrLen(IPAddress) > 15)
		Return 0

	IfInString, IPAddress, %A_Space%
		Return 0

	StringSplit, Octets, IPAddress, .
	if (Octets0 <> 4)
		Return 0
	Loop 4
	{
		If Octets%A_Index% is not digit
			Return 0
		If (Octets%A_Index% < 0 or Octets%A_Index% > 255)
			Return 0
	}

	return 1
}

if (CusBrand = "" || CusBrand = ) {
 CusBrand := "A valid printer BRAND" "`n"
 errmsg := CusBrand
}

if (CusName = "" || CusName = ) {
 CusName := "A valid printer NAME" "`n"
 errmsg := errmsg CusName
}

if (ValidIP(CusIP) = 0) {
 CusIP := "A valid IP ADDRESS" "`n"
 errmsg := errmsg CusIP
}

if (errmsg != "" || errmsg != ) {
	MsgBox,16,%A_ScriptName%, % "Please select or enter the following: " "`n" "`n" errmsg
	return
	}
	else {
		If (CusBrand = "HP"){
			printerName := CusName , ipADDRESS := CusIP , driverName := HpUDName , driverPath:= HpPath
			installPrinter()
		}
		else If (CusBrand = "RICOH/GESTETNER"){
			printerName := CusName , ipADDRESS := CusIP , driverName := RicohUDName , driverPath:= RicohPath
			installPrinter()
		}
		else If (CusBrand = "SHARP"){
			printerName := CusName , ipADDRESS := CusIP , driverName := SharpUDName , driverPath:= SharpPath
			installPrinter()
		}
		else If (CusBrand = "XEROX"){
			printerName := CusName , ipADDRESS := CusIP , driverName := XeroxUDName , driverPath:= XeroxPath
			installPrinter()
		}
		else If (CusBrand = "LEXMARK"){
			printerName := CusName , ipADDRESS := CusIP , driverName := LexmarkUDName , driverPath:= LexmarkPath
			installPrinter()
		}
		else If (CusBrand = "KYOCERA"){
			printerName := CusName , ipADDRESS := CusIP , driverName := KyoceraUDName , driverPath:= KyoceraPath
			installPrinter()
		}
		else If (CusBrand = "SAMSUNG"){
			printerName := CusName , ipADDRESS := CusIP , driverName := SamsungUDName , driverPath:= SamsungPath
			installPrinter()
		}
	}
}

return

GuiClose:
ExitApp

!x::
ExitApp
return