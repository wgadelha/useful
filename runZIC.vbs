'#==============================================================================
'#==============================================================================
'#  SCRIPT.........:  runZIC.vbs
'#  AUTHOR.........:  Walter Gadelha
'#  VERSION........:  0.1alpha
'#  CREATED........:  17/12/2020
'#  REQUIREMENTS...:  Windows 7 / 10
'#  DESCRIPTION....:  Opens an URL after disabling IE Automatic configuration script
'#
'#  NOTES..........:  Script based on Enable Disable IE Proxy by Stuart Barret
'#                    Usage of MsgBox for fast deployment
'#                    IP addresses changed for privacy reasons
'#
'#==============================================================================
'#==============================================================================

Const HKEY_CURRENT_USER = &H80000001
Set objShell = CreateObject("WScript.Shell")
strPC = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

On Error Resume Next

Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

'#==============================================================================
'#  Retrieve Binary value from the keypath:
'#	[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections]
'#  "DefaultConnectionSettings"=
'#  hex:3c,00,00,00,1f,00,00,00,09,00,00,00,00,[...]9e,00,00,00,00,00,00,00,00
'#  
'#  The 9th byte of the key defines IE LAN Settings options as a bitfield
'#  
'#  * 0x1: (Always 1)
'#  * 0x2: Proxy enabled
'#  * 0x4: "Use automatic configuration script" checked
'#  * 0x8: "Automatically detect settings" checked
'#
'#==============================================================================

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
strValueName = "DefaultConnectionSettings"

objReg.GetBinaryValue HKEY_CURRENT_USER,strKeyPath,strValueName,bValue

'#==============================================================================
'#  Asks if the user wants to open ZIC, if the user select YES 
'#  Disable IE Automatic configuration script option then ask user 
'#  which ZIC version (R6 or R5) will be used to open the default browser
'#	
'#  If the user does not want to open ZIC, it can click NO to enable
'#  IE Automatic configuration script, or CANCEL to quit the script without
'#  any changes to the IE configuration
'#  
'#==============================================================================

strR6Addr = "http://a.b.c.d:8080/index.php"
strR5Addr = "http://a.b.c.d:8080//"

IEPrompt = MsgBox ("Do you want to open ZIC R6 or R5?"  & vbCrLf & vbCrLf & "Yes will DISABLE IE Automatic configuration script."  & vbCrLf & "No will ENABLE IE Automatic configuration script.",vbQuestion+vbYesNoCancel+vbDefaultButton1, "Run ZIC")

If IEPrompt = vbYes Then
	bValue(8) = bValue(8) And Not 4
	objReg.SetBinaryValue HKEY_CURRENT_USER,strKeyPath,strValueName,bValue 

	ZICPrompt = MsgBox ("Do you want to open R6?" & vbCrLf & vbCrLf & "Yes for R6." & vbCrLf & "No for R5.", vbQuestion+vbYesNoCancel+vbDefaultButton1, "Run ZIC")
	If ZICPrompt = vbYes Then
		objShell.Run strR6Addr
		
		ElseIf ZICPrompt = vbNo Then
		objShell.Run strR5Addr
	End If

	ElseIf IEPrompt = vbNo Then
		bValue(8) = bValue(8) Or 4
		objReg.SetBinaryValue HKEY_CURRENT_USER,strKeyPath,strValueName,bValue 
		MsgBox "IE Automatic configuration script is now ENABLED on " & strPC & _
		".",vbInformation+vbDefaultButton1, "Run ZIC"
End If
