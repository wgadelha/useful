'#==============================================================================
'#==============================================================================
'#  SCRIPT.........:	Enable Disable IE Proxy.vbs
'#  AUTHOR.........:	Stuart Barrett
'#  VERSION........:	2.0
'#  CREATED........:	15/09/10
'#  COPYRIGHT......:	2010
'#  LICENSE........:	Freeware
'#  REQUIREMENTS...:  
'#  DESCRIPTION....:	Enables or disables IE proxy on the local system
'#
'#  NOTES..........:  
'# 
'#  CUSTOMIZE......:  
'#==============================================================================
'#  REVISED BY.....:	Stuart Barrett
'#  EMAIL..........:  
'#  REVISION DATE..:	27/02/15
'#  REVISION NOTES.:	Added in ping on beginning of script to force quit if
'#						required
'#
'#==============================================================================
'#==============================================================================

'#==============================================================================
'#	Change this variable to a machine name / IP on your network
'#==============================================================================
	strRemoteMachine = "192.168.254.253"

Const HKEY_CURRENT_USER = &H80000001
Set objShell = CreateObject("WScript.Shell")
strPC = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strDomain = objShell.ExpandEnvironmentStrings("%USERDOMAIN%")

On Error Resume Next

If Reachable(strRemoteMachine) Then
	MsgBox strPC & " is currently connected to the " & strDomain & _
		" domain." & vbCrLF & vbCrLF & " This script will now exit.", vbExclamation, "Enable / Disable IE Proxy"
	WScript.Quit
End If

Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
strValueName = "ProxyEnable"

objReg.GetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

'#==============================================================================
'#	If IE Proxy is currently enabled display message and ask user whether it
'#	should then be disabled
'#==============================================================================
If dwValue = 1 Then
	IEPrompt = MsgBox ("IE Proxy is currently ENABLED on " & strPC & _
	".  Do you want to DISABLE it?", vbQuestion+vbYesNo, "Disable IE Proxy")
	If IEPrompt = vbYes Then
		dwValue = 0
		objReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue 
		MsgBox "IE Proxy is now DISABLED on " & strPC & _
		".",vbInformation, "Disable IE Proxy"
		ElseIf IEPrompt = vbNo Then
			MsgBox "IE Proxy is still ENABLED on " & strPC & _
			".",vbInformation, "Disable IE Proxy"
	End If
	'#==============================================================================
	'#	If IE Proxy is currently disabled display message and ask user whether it
	'#	should then be enabled
	'#==============================================================================
	ElseIf dwValue = 0 Then
		IEPrompt = MsgBox ("IE Proxy is currently DISABLED on " & strPC & _
		".  Do you want to ENABLE it?", vbQuestion+vbYesNo, "Enable IE Proxy")
		If IEPrompt = vbYes Then
			dwValue = 1
			objReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
			MsgBox "IE Proxy is now ENABLED on " & strPC & _
			".",vbInformation, "Enable IE Proxy"
			ElseIf IEPrompt = vbNo Then
				MsgBox "IE Proxy is still DISABLED on " & strPC & _
				".",vbInformation, "Enable IE Proxy"
	End If
End If

strKeyPath = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
strValueName = "Proxy"
dwValue = 1

objReg.CreateKey HKEY_CURRENT_USER,strKeyPath
objReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

'#--------------------------------------------------------------------------
'#  FUNCTION.......:	Reachable(strComp)
'#  PURPOSE........:	Checks whether the remote machine is online
'#  ARGUMENTS......:	
'#  EXAMPLE........:	Reachable(PC1)
'#  NOTES..........:  
'#--------------------------------------------------------------------------
Function Reachable(strComp)
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colPing = objWMIService.ExecQuery _
		("Select * from Win32_PingStatus Where Address = '" & strComp & "'")
	For Each objItemR in colPing
		If IsNull(objItemR.StatusCode) Or objItemR.StatusCode <> 0 Then
			Reachable = False
			Else
				Reachable = True
		End If
	Next
End Function