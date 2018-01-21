On Error Resume Next

WScript.Quit(VBSMain)

Function VBSMain()
	If CheckOSVersion() <= 0 Then
		WScript.Echo "Only for Windows XP and earlier casue later Windows already supported this feature natively."
		WScript.Quit
	End If

	If WScript.Arguments.Count = 0 Then
		Call RegLnkShellCmd(1)
	Else
		If WScript.Arguments(0) = "u" Then
			Call RegLnkShellCmd(0)
		Else
			Call FindShortCutTarget(WScript.Arguments(0))
		End If
	End If
End Function

Sub RegLnkShellCmd(pOnOff)
	Set Shell = CreateObject("WScript.Shell")
	If pOnOff = 1 Then
		Shell.RegWrite "HKCR\lnkfile\shell\Find Target\command\", "WScript.exe " & WScript.ScriptFullName & " ""%l"""
	Else
		Shell.RegDelete "HKCR\lnkfile\shell\Find Target\command\"
		Shell.RegDelete "HKCR\lnkfile\shell\Find Target\"
	End If
End Sub

Sub FindShortCutTarget(pShortCut)
	Set Shell = CreateObject("WScript.Shell")
	Set link = Shell.CreateShortcut(pShortCut)
	Shell.Run "explorer /select," & link.TargetPath
End Sub

Function CheckOSVersion()
On Error Resume Next
	osxp = "5.1"
	Set WSHShell = WScript.CreateObject("WScript.Shell") 
	strRegKey = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentVersion" 
	osVersion = WSHShell.RegRead(strRegKey)
	If Err.Number <> 0 Then
		Err.Clear
		CheckOSVersion = 0
		Exit Function
	End If
	If (osVersion <= osxp) Then CheckOSVersion = 1 Else CheckOSVersion = 0
End Function