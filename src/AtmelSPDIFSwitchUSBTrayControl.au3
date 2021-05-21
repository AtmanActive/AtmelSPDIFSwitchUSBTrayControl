#include <Array.au3>
#include <AutoItConstants.au3>
#include <CommMG64.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GuiConstantsEx.au3>
#include <GuiEdit.au3>
#include <GuiListBox.au3>
#include <Misc.au3>
#include <NamedPipes.au3>
#include <Process.au3>
#include <ScrollBarsConstants.au3>
#include <StaticConstants.au3>
#include <TrayConstants.au3>
#include <WinAPI.au3>
#include <WinAPIProc.au3>
#include <WindowsConstants.au3>



; SINGLE INSTANCE ONLY
If _Singleton( "AtmelSPDIFSwitchUSBTrayControl", 1 ) = 0 Then
	Exit
EndIf







; ===============================================================================================================================
; Global constants
; ===============================================================================================================================

Global Const $PROGNAME = "Atmel S/PDIF-Switch USB Tray Control"

Global Const $BUFSIZE = 4096
Global Const $DEFCMD = "cmd.exe /c dir c:\windows /s"
Global Const $PIPE_NAME = "\\.\\pipe\\APipeName"
Global Const $ERROR_MORE_DATA = 234

Global Const $FILE_FLAG_OVERLAPPED = 0x40000000
Global Const $FILE_FLAG_NO_BUFFERING = 0x20000000

Global $logcharlimit  = 15000 ; Max size of edit
Global $logcharbuffer = 500  ; Buffer to hold 2 x longest expected line

; ===============================================================================================================================
; Global variables
; ===============================================================================================================================

Global $gui_window, $gui_window_is_hidden

Global $iEdit, $iMemo, $iSend, $iServer, $hPipe, $fork_pid, $fork_cmd_pid, $forked_pids_array

Global $tray_event, $tray_menu_item_show, $tray_menu_item_exit, $tray_menu_item_options, $tray_menu_item_cla1, $tray_menu_item_cla0, $tray_menu_item_cli1, $tray_menu_item_cli0, $tray_menu_item_clr1, $tray_menu_item_clr0, $tray_menu_item_clm0, $tray_menu_item_clm1, $tray_menu_item_ci1, $tray_menu_item_ci0, $tray_menu_item_cz1, $tray_menu_item_cz0

Global $tray_menu_channel[8]
Global $channel_name[8]
Global $channel_enabled[8]
Global $channel_icon_present[8]
Global $channel_icon_name[8]
Global $com_port_name, $com_port_number
Global $autorun_channel, $autorun_icon, $exitrun_channel

Global $my_exe_path = @ScriptFullPath
Dim $szDrive, $szDir, $szFName, $szExt
$TestPath = _PathSplit( $my_exe_path, $szDrive, $szDir, $szFName, $szExt )
Global $my_exe_dir = $szDrive & $szDir
Global $my_ini_path = $my_exe_dir & "\AtmelSPDIFSwitchUSBTrayControl.ini"

; ===============================================================================================================================
; Main
; ===============================================================================================================================

EnvSet( "HOME",   $my_exe_dir )
EnvSet( "PATH",   $my_exe_dir )



ReadINICOM()
CreateGUIWindow()
print( $PROGNAME & " v1.2 " & "booting up..." )
OpenCOM(1)
ReadINIChannels()
CreateGUITray()
SetAutorun()
Set_Detection_Timeout(1)
Set_Option( "Disabled invalid packets detection", "$CI0" )
Set_Option( "Disabled zero packets detection", "$CZ0" )
_CommClosePort()
CPU_Loop()








Func ReadINICOM()
	
	Local $com = IniRead ( $my_ini_path, "Com", "Port", "0" )
	If $com = "0" Then
		MsgBox( 0, $PROGNAME, "This program needs section [Com], key 'Port' to be set to a valid com port name in this program's ini file. The program will now exit." )
		clean_up()
	Else
		$com_port_name = $com
	EndIf
	
	Local $nbr = StringRegExp( $com_port_name, '\d+$', 1 )
	If IsArray( $nbr ) Then ; Msgbox( 0,"", $nbr[0] )
		$com_port_number = $nbr[0]
	Else
		MsgBox( 0, $PROGNAME, "This program needs to extract COM port number from supplied COM port name. The value in the ini file '" & $com_port_name & "' does not yield any number. The program will now exit." )
		clean_up()
	EndIf
	
EndFunc   ;==> ReadINICOM





Func ReadINIChannels()
	
	For $i = 1 To 7
		Local $name = IniRead ( $my_ini_path, "Channel Names", $i, "0" )
		If $name = "0" Then
			$channel_enabled[$i] = 0
		Else
			$channel_name[$i] = $name
			$channel_enabled[$i] = 1
			print( "Channel " & $i & ": " & $name )
		EndIf
	Next
	
	For $i = 1 To 7
		Local $icon = IniRead ( $my_ini_path, "Channel Icons", $i, "0" )
		;print( $i & " - " & $icon )
		If $icon = "0" Then
			$channel_icon_present[$i] = 0
		Else
			$channel_icon_name[$i] = $icon
			$channel_icon_present[$i] = 1
		EndIf
	Next
	
	$autorun_channel = IniRead ( $my_ini_path, "OnRun Set", "SetChannel", "0" )
	$autorun_icon    = IniRead ( $my_ini_path, "OnRun Set", "Icon", "0" )
	$exitrun_channel = IniRead ( $my_ini_path, "OnExit Set", "SetChannel", "0" )
	
EndFunc   ;==> ReadINIChannels





Func OpenCOM( $silent_or_verbose )
	
	Local $PortsArray = _CommListPorts( 0 )
	
	if $PortsArray = "" Then
		MsgBox( 0, $PROGNAME, "No COM ports found on this computer. Did you plug-in your device? The program will now exit." )
		clean_up()
	EndIf
	
	Local $portmatch = 0
	
	if $silent_or_verbose = 1 Then
		print( "Found the following COM ports:" )
	EndIf
	For $i = 1 To $PortsArray[0]
		if $silent_or_verbose = 1 Then
			print( $PortsArray[$i] )
		EndIf
		if $com_port_name = $PortsArray[$i] Then
			$portmatch = 1
		EndIf
	Next
	
	if $portmatch = 0 Then
		print( "Can't match to expected " & $com_port_name )
		MsgBox( 0, $PROGNAME, "No COM ports found on this computer match the one specified in the ini file: " & $com_port_name & ". Please check the name and fix the ini file. The program will now exit." )
		clean_up()
	EndIf
	
	Local $sErr
	Local $portOpenOK = _CommSetport( $com_port_number, $sErr, 9600, 8, 0, 1, 0, 0, 0 )
	if $portOpenOK = 1 Then
		if $silent_or_verbose = 1 Then
			print( "COM port " & $com_port_name & " opened successfully." )
		EndIf
	Else
		print( "Can't open COM port " & $com_port_name )
		MsgBox( 0, $PROGNAME, "Can't open COM port " & $com_port_name & ". Please check the permissions and if some other app is occupying it. The program will now exit." )
		clean_up()
	EndIf
	
EndFunc   ;==> OpenCOM





Func CreateGUIWindow()
	
	$gui_window = GUICreate( $PROGNAME, 1000, 500, -1, -1, -1 )
	GUISetBkColor( 0x000000 )
	$iMemo = GUICtrlCreateEdit("", 0, 0, _WinAPI_GetClientWidth( $gui_window ), 500, $ES_MULTILINE + $ES_AUTOVSCROLL + $WS_VSCROLL )
	GUICtrlSetFont( $iMemo, 10, 400, 0, "Consolas" )
	GUICtrlSetBkColor( -1, 0x001010 )
	GUICtrlSetColor( -1, 0x00FFFF )
	_GUICtrlEdit_SetLimitText( $iMemo, $logcharlimit )
	GUISetState()
	GUISetIcon ( $my_exe_dir & "\" & "AtmelSPDIFSwitchUSBTrayControl.icl", 1, $gui_window )
	
EndFunc   ;==> CreateGUIWindow





Func CreateGUITray()
	
	populate_random_menu_ids()
	
	Opt( "TrayOnEventMode", 0 ) ; Enable/disable OnEvent functions notifications for the tray.
	Opt( "TrayMenuMode", 3 ) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
	TraySetIcon( $my_exe_dir & "\" & "AtmelSPDIFSwitchUSBTrayControl.icl", 1 )
	TraySetClick( 8 ) ;only show the menu when right clicking
	
	For $i = 1 To 7
		If $channel_enabled[$i] = 1 Then
			$tray_menu_channel[$i] = TrayCreateItem( $i & ": " & $channel_name[$i], -1, -1, $TRAY_ITEM_RADIO )
		EndIf
	Next
	
	TrayCreateItem( "" ) ; separator
	
	$tray_menu_item_options = TrayCreateMenu( "Options" )
	
	$tray_menu_item_cla1 = TrayCreateItem( "Enable LEDs on current active input (solid LED)", $tray_menu_item_options )
	$tray_menu_item_cla0 = TrayCreateItem( "Disable LEDs on current active input (solid LED)", $tray_menu_item_options )
	
	TrayCreateItem( "", $tray_menu_item_options ) ; separator
	
	$tray_menu_item_cli1 = TrayCreateItem( "Enable LEDs on inactive inputs (blinking LEDs)", $tray_menu_item_options )
	$tray_menu_item_cli0 = TrayCreateItem( "Disable LEDs on inactive inputs (blinking LEDs)", $tray_menu_item_options )
	
	TrayCreateItem( "", $tray_menu_item_options ) ; separator
	
	$tray_menu_item_clr1 = TrayCreateItem( "Enable LEDs on inactive inputs without signal (only blinking red LEDs)", $tray_menu_item_options )
	$tray_menu_item_clr0 = TrayCreateItem( "Disable LEDs on inactive inputs without signal (only blinking red LEDs)", $tray_menu_item_options )
	
	TrayCreateItem( "", $tray_menu_item_options ) ; separator
	
	$tray_menu_item_clm0 = TrayCreateItem( "Enable Alternative LED Mode", $tray_menu_item_options )
	$tray_menu_item_clm1 = TrayCreateItem( "Disable Alternative LED Mode", $tray_menu_item_options )
	
	TrayCreateItem( "", $tray_menu_item_options ) ; separator
	
	$tray_menu_item_ci1  = TrayCreateItem( "Enable invalid packets detection", $tray_menu_item_options )
	$tray_menu_item_ci0  = TrayCreateItem( "Disable invalid packets detection", $tray_menu_item_options )
	
	TrayCreateItem( "", $tray_menu_item_options ) ; separator
	
	$tray_menu_item_cz1  = TrayCreateItem( "Enable zero packets detection", $tray_menu_item_options )
	$tray_menu_item_cz0  = TrayCreateItem( "Disable zero packets detection", $tray_menu_item_options )
	
	TrayCreateItem( "" ) ; separator
	
	$tray_menu_item_show = TrayCreateItem( "Show Log" )
	$tray_menu_item_exit = TrayCreateItem( "Exit" )
	
	TraySetState( $TRAY_ICONSTATE_SHOW ) ; Show the tray menu.
	TraySetToolTip( $PROGNAME )
	
	GUISetState( @SW_HIDE, $gui_window )
	$gui_window_is_hidden = 1
	
EndFunc   ;==> CreateGUITray





Func SetAutorun()
	
	If $autorun_icon = "0" Then
	Else
		TraySetIcon( $my_exe_dir & "\" & "ico\" & $autorun_icon )
		GUISetIcon ( $my_exe_dir & "\" & "ico\" & $autorun_icon )
	EndIf
	
	If $autorun_channel = "0" Then
	Else
		If $autorun_channel > 0 And $autorun_channel < 8 Then
			Set_Channel_Number( $autorun_channel )
			TrayItemSetState( $tray_menu_channel[ $autorun_channel ], $TRAY_CHECKED )
		EndIf
	EndIf
	
EndFunc   ;==> SetAutorun()





Func GUI_Loop()
	
	Switch GUIGetMsg()
		Case $GUI_EVENT_MINIMIZE
			GUISetState( @SW_HIDE, $gui_window )
			$gui_window_is_hidden = 1
		Case $GUI_EVENT_CLOSE
			clean_up()
	EndSwitch
	
	Switch TrayGetMsg()
		Case $TRAY_EVENT_PRIMARYDOWN
			If $gui_window_is_hidden = 1 Then
				GUISetState( @SW_SHOW,    $gui_window )
				GUISetState( @SW_RESTORE, $gui_window )
				$gui_window_is_hidden = 0
			Else
				GUISetState( @SW_HIDE, $gui_window )
				$gui_window_is_hidden = 1
			EndIf
		Case $tray_menu_item_show
			If $gui_window_is_hidden = 1 Then
				GUISetState( @SW_SHOW,    $gui_window )
				GUISetState( @SW_RESTORE, $gui_window )
				$gui_window_is_hidden = 0
			Else
				GUISetState( @SW_HIDE, $gui_window )
				$gui_window_is_hidden = 1
			EndIf
		Case $tray_menu_item_exit ; Exit the loop.
			clean_up()
		Case $tray_menu_channel[1]
			If $channel_enabled[1] = 1 Then
				OpenCOM(0)
				Set_Channel_Number(1)
				_CommClosePort()
			EndIf
		Case $tray_menu_channel[2]
			If $channel_enabled[2] = 1 Then
				OpenCOM(0)
				Set_Channel_Number(2)
				_CommClosePort()
			EndIf
		Case $tray_menu_channel[3]
			If $channel_enabled[3] = 1 Then
				OpenCOM(0)
				Set_Channel_Number(3)
				_CommClosePort()
			EndIf
		Case $tray_menu_channel[4]
			If $channel_enabled[4] = 1 Then
				OpenCOM(0)
				Set_Channel_Number(4)
				_CommClosePort()
			EndIf
		Case $tray_menu_channel[5]
			If $channel_enabled[5] = 1 Then
				OpenCOM(0)
				Set_Channel_Number(5)
				_CommClosePort()
			EndIf
		Case $tray_menu_channel[6]
			If $channel_enabled[6] = 1 Then
				OpenCOM(0)
				Set_Channel_Number(6)
				_CommClosePort()
			EndIf
		Case $tray_menu_channel[7]
			If $channel_enabled[7] = 1 Then
				OpenCOM(0)
				Set_Channel_Number(7)
				_CommClosePort()
			EndIf
		Case $tray_menu_item_options, $tray_menu_item_cla1
			OpenCOM(0)
			Set_Option( "Enabled LEDs on current active input (solid LED)", "$CLA1" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_cla0
			OpenCOM(0)
			Set_Option( "Disabled LEDs on current active input (solid LED)", "$CLA0" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_cli0
			OpenCOM(0)
			Set_Option( "Disabled LEDs on inactive inputs (blinking LEDs)", "$CLI0" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_cli1
			OpenCOM(0)
			Set_Option( "Enabled LEDs on inactive inputs (blinking LEDs)", "$CLI1" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_clr1
			OpenCOM(0)
			Set_Option( "Enabled LEDs on inactive inputs without signal (only blinking red LEDs)", "$CLR1" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_clr0
			OpenCOM(0)
			Set_Option( "Disabled LEDs on inactive inputs without signal (only blinking red LEDs)", "$CLR0" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_clm0
			OpenCOM(0)
			Set_Option( "Enabled Alternative LED Mode", "$CLM0" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_clm1
			OpenCOM(0)
			Set_Option( "Disable Alternative LED Mode", "$CLM1" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_ci0
			OpenCOM(0)
			Set_Option( "Disabled invalid packets detection", "$CI0" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_ci1
			OpenCOM(0)
			Set_Option( "Enabled invalid packets detection", "$CI1" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_cz0
			OpenCOM(0)
			Set_Option( "Disabled zero packets detection", "$CZ0" )
			_CommClosePort()
		Case $tray_menu_item_options, $tray_menu_item_cz1
			OpenCOM(0)
			Set_Option( "Enable zero packets detection", "$CZ1" )
			_CommClosePort()
	EndSwitch
	
EndFunc ;==> GUI_Loop()





Func CPU_Loop()
	While 1
		GUI_Loop()
		sleep ( 10 ) ; 100Hz
	WEnd
EndFunc ;==> CPU_Loop()





Func Set_Channel_Number( $number )
	Faster_Switching()
	print( "Setting switch output to channel number " & $number & ": " & $channel_name[$number] )
	_CommSendString( "$I" & $number & @CRLF, 0 )
	If $channel_icon_present[$number] = 1 Then
		TraySetIcon( $my_exe_dir & "\" & "ico\" & $channel_icon_name[$number] )
	Else
		TraySetIcon( $my_exe_dir & "\" & "AtmelSPDIFSwitchUSBTrayControl.icl", ( 1 + $number ) )
	EndIf
EndFunc ;==> Set_Channel_Number( $number )




Func Set_Detection_Timeout( $number )
	print( "Setting detection timeout to " & $number & " seconds" )
	_CommSendString( "$CT" & $number & @CRLF, 0 )
EndFunc ;==> Set_Detection_Timeout( $number )



Func Faster_Switching()
	_CommSendString( "$CT1" & @CRLF, 0 )
	_CommSendString( "$CI0" & @CRLF, 0 )
	_CommSendString( "$CZ0" & @CRLF, 0 )
EndFunc ;==> Faster_Switching()





Func Set_Option( $human_string, $serial_string )
	print( $human_string )
	_CommSendString( $serial_string & @CRLF, 0 )
EndFunc ;==> Set_Channel_Number( $number )





Func print( $msg )
	
	_GUICtrlEdit_BeginUpdate( $iMemo )
	
	; Get line count
	$iLines_Count = _GUICtrlEdit_GetLineCount( $iMemo )
	
	; Check position of first char of last line is not inside buffer
	If _GUICtrlEdit_LineIndex( $iMemo, $iLines_Count - 1 ) > $logcharlimit - $logcharbuffer Then
		; Move down lines until we get to a value greater then the buffer
		For $i = 1 To $iLines_Count - 1
			$iCurrLine_Index = _GUICtrlEdit_LineIndex( $iMemo, $i )
			If $iCurrLine_Index > $logcharbuffer Then
				; Select and delete lines to that point
				_GUICtrlEdit_SetSel( $iMemo, 0, $iCurrLine_Index - 1 )
				_GUICtrlEdit_ReplaceSel( $iMemo, "", False )
				; No point in looking further
				ExitLoop
			EndIf
		Next
	EndIf
	
	; Text is now smaller by at least twice the buffer size so add new text
	_GUICtrlEdit_AppendText( $iMemo, $msg & @CRLF )
	_GUICtrlEdit_EndUpdate( $iMemo )
	; Select the last character and scroll it into view
	_GUICtrlEdit_SetSel( $iMemo, -1, -1 )
	_GUICtrlEdit_Scroll( $iMemo, $SB_SCROLLCARET )
	
EndFunc ;==> print



Func populate_random_menu_ids()
	For $i = 1 To 7
		$tray_menu_channel[$i] = generate_random_string($i)
		;print( $tray_menu_channel[$i] )
	Next
EndFunc   ;==> generate_random_string( $slot )


Func generate_random_string( $slot )
    Local $sText = ""
    For $i = 1 To 32
        $sText &= Chr( Random( 65, 122, 1 ) ) ; Return an integer between 65 and 122 which represent the ASCII characters between a (lower-case) to Z (upper-case).
    Next
		Return $sText
EndFunc   ;==> generate_random_string( $slot )





Func clean_up()
	
	If $exitrun_channel = "0" Then
	Else
		If $exitrun_channel > 0 And $exitrun_channel < 8 Then
			OpenCOM(0)
			Set_Channel_Number( $exitrun_channel )
			_CommClosePort()
		EndIf
	EndIf
	;_CommClosePort()
	exit

EndFunc ; ==> clean_up()
