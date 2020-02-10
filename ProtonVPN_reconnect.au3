#include <Array.au3>
#include <GuiToolBar.au3>
#include <WinAPIDiag.au3>
#include <Array.au3>

Global $hSysTray_Handle

$vpnWindow = "[TITLE:ProtonVPN; REGEXPCLASS:ProtonVPN\.exe]"
$program = "C:\Program Files (x86)\Proton Technologies\ProtonVPN\ProtonVPN.exe"



$connected = False
$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$oWMIService = ObjGet("winmgmts:\\localhost\ROOT\CIMV2")
$colItems = $oWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter Where Name like '%ProtonVPN%'", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

If IsObj($colItems) Then
   For $objItem In $colItems
      If $objItem.NetConnectionStatus = 2 Then
		 $connected = True
		 ExitLoop
	  EndIf
   Next
EndIf

If Not WinExists($vpnWindow) Then
   Run($program)
   WinWait($vpnWindow)
   Sleep(5000)
   ;ConsoleWrite('done')
EndIf



Click_SysTray_Icon('', $vpnWindow)
CloseAd()

WinActivate($vpnWindow)
WinWaitActive($vpnWindow)
$aPos = WinGetPos($vpnWindow)
If $connected Then
   _MouseClickFast("", $aPos[0]+140, $aPos[1]+240)
   ;Sleep(100)
Else
   _MouseClickFast("", $aPos[0]+350, $aPos[1]+400)
   ;Sleep(50)
   ;_MouseClickFast("", $aPos[0]+140, $aPos[1]+240)
   _MouseClickFast("", $aPos[0]+250, $aPos[1]+310)
   MouseClick("", $aPos[0]+265, $aPos[1]+420)
EndIf


Sleep(2000)
ConsoleWrite("VPN (re)connected"&@CRLF)



Func CloseAd()
   $aList=WinList($vpnWindow, "")
   If $aList[0][0]=3 Then
	  WinActivate($aList[1][1])
	  $pos = WinGetPos($aList[1][1])
	  ConsoleWrite($pos[0]&'-'&$pos[1]&'-'&$pos[2]&'-'&$pos[3]&@CR)
	  MouseClick("", $pos[0]+$pos[2]-5, $pos[1]+5)
   EndIf
EndFunc

Func Click_SysTray_Icon($sSearch, $title)
   ;ConsoleWrite('test'&$hSysTray_Handle&@CR)
   $hSysTray_Handle = ControlGetHandle('[Class:Shell_TrayWnd]', '', '[Class:ToolbarWindow32;Instance:1]')
   ;ConsoleWrite('test1 '&$hSysTray_Handle&@CR)
   $iButton=Get_SysTray_IconText($sSearch)
   For $iSysTray_ButtonNumber In $iButton
	  _GUICtrlToolbar_ClickButton($hSysTray_Handle, $iSysTray_ButtonNumber, "left", False, 2)
	  If BitAND(WinGetState($vpnWindow), $WIN_STATE_VISIBLE) Then Return
   Next

   $hSysTray_Handle = ControlGetHandle('[Class:NotifyIconOverflowWindow]', '', '[Class:ToolbarWindow32;Instance:1]')
   ;ConsoleWrite('test2 '&$hSysTray_Handle&@CR)
   $iButton=Get_SysTray_IconText($sSearch)
   For $iSysTray_ButtonNumber In $iButton
	  ControlClick(WinGetHandle("[Class:Shell_TrayWnd]", ""), "", "[CLASS:Button; INSTANCE:2]")
	  _GUICtrlToolbar_ClickButton($hSysTray_Handle, $iSysTray_ButtonNumber, "left", False, 2)
	  If BitAND(WinGetState($vpnWindow), $WIN_STATE_VISIBLE) Then Return
   Next
EndFunc

Func Get_SysTray_IconText($sSearch)
   ; Get systray item count
   Local $iSysTray_ButCount = _GUICtrlToolbar_ButtonCount($hSysTray_Handle)
   Local $aSysTray_Match[0]

   ; Look for wanted tooltip
   For $iSysTray_ButtonNumber = 0 To $iSysTray_ButCount - 1
	  ;ConsoleWrite(@CR & $iSysTray_ButtonNumber & ":" & _GUICtrlToolbar_GetButtonText($hSysTray_Handle, $iSysTray_ButtonNumber))
	  If $sSearch= _GUICtrlToolbar_GetButtonText($hSysTray_Handle, $iSysTray_ButtonNumber) Then _
		 _ArrayAdd($aSysTray_Match, $iSysTray_ButtonNumber)
   Next
   ;_ArrayDisplay($aSysTray_Match)
   Return $aSysTray_Match
EndFunc   ;==>Get_SysTray_IconText

;~ Func Get_SysTray_IconText($sSearch)
;~     For $i = 1 To 99
;~         ; Get systray item count
;~         Local $iSysTray_ButCount = _GUICtrlToolbar_ButtonCount($hSysTray_Handle)
;~         If $iSysTray_ButCount = 0 Then
;~             ;MsgBox(16, "Error", "No items found in system tray")
;~             ContinueLoop
;~         EndIf

;~         Local $aSysTray_ButtonText[$iSysTray_ButCount]

;~         ; Look for wanted tooltip
;~         For $iSysTray_ButtonNumber = 0 To $iSysTray_ButCount - 1
;~ 		    ;ConsoleWrite(@CR & $i & " " & $iSysTray_ButtonNumber & ":" & _GUICtrlToolbar_GetButtonText($hSysTray_Handle, $iSysTray_ButtonNumber))
;~             If $sSearch= _GUICtrlToolbar_GetButtonText($hSysTray_Handle, $iSysTray_ButtonNumber) Then _
;~             Return SetError(0, $i, $iSysTray_ButtonNumber)
;~         Next
;~     Next
;~     Return SetError(1, -1, -1)

;~ EndFunc   ;==>Get_SysTray_IconText


Func _MouseClickFast($w, $x, $y)
    $x = $x*65535/@DesktopWidth
    $y = $y*65535/@DesktopHeight

    _WinAPI_Mouse_Event(BitOR($MOUSEEVENTF_ABSOLUTE, $MOUSEEVENTF_MOVE), $x, $y)
    _WinAPI_Mouse_Event(BitOR($MOUSEEVENTF_ABSOLUTE, $MOUSEEVENTF_LEFTDOWN), $x, $y)
    _WinAPI_Mouse_Event(BitOR($MOUSEEVENTF_ABSOLUTE, $MOUSEEVENTF_LEFTUP), $x, $y)
EndFunc

Func _MouseMoveFast($w, $x, $y)
    $x = $x*65535/@DesktopWidth
    $y = $y*65535/@DesktopHeight

    _WinAPI_Mouse_Event(BitOR($MOUSEEVENTF_ABSOLUTE, $MOUSEEVENTF_MOVE), $x, $y)
EndFunc