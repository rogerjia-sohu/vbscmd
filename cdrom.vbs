'Script Format:
'	cdrom.vbs lDevNum strAction
'	(e.g.: cdrom 1 eject)
'Composed by Roger 2007/12/23

With WScript .Quit(VBSMain(.Arguments.Count, .Arguments)) End With

Function VBSMain(argc, argv)
On Error Resume Next
	If argc = 2 Then
		VBSMain = MsftDisc(argv(0), argv(1))
	Else
		VBSMain = cdromUsage()
	End If
	DisplayErrorInfo
End Function

Sub DisplayErrorInfo
	If Err <= 0 Then Exit Sub
	ErrInfo = "Error:" & vbTab & Err & vbTab & "hex: " & "&H" & Hex(Err) & vbNewLine _
		& "Source:" & vbTab & Err.Source & vbNewLine _
		& "Desc.:" & vbTab & Err.Description
	WScript.Echo ErrInfo
	Err.Clear
End Sub

Function cdromUsage()
	strMsg = "Usage: cdrom devnum oper" & vbNewLine _
		& vbTab & "devnum" & vbTab & "Device number, 0(zero) for all. One value at a time!" & vbNewLine  _
		& vbTab & "oper" & vbTab & "Operator: {eject|close}" & vbNewLine
	WScript.Echo strMsg
	cdromUsage = 0
End Function

Function MsftDisc(lDevNum, strAction)
	strOper = LCase(Trim(strAction))
	If Not (strOper = "eject" Or strOper = "close") Then
		MsftDisc = 0
		WScript.Echo "Unknown Action:" & vbTab & strAction
		Exit Function
	End If

	Num = 0
	IsAll = CLng(lDevNum) <= 0
	Set colDiscMaster = CreateObject("IMAPI2.MsftDiscMaster2")
	For Each Id In colDiscMaster
		Num = Num + 1
		If IsAll Or CLng(Num) = CLng(lDevNum) Then
			Set objRecorder = CreateObject("IMAPI2.MsftDiscRecorder2")
			objRecorder.InitializeDiscRecorder Id
			Select Case strOper
			Case "eject"
				objRecorder.EjectMedia()
			Case "close"
				objRecorder.CloseTray()
			End Select
		End If
	Next
End Function