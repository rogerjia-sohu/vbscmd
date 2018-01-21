'----------------------------------------------------------------------
'
' Copyright (c) JEDAT Inc. All rights reserved.
'
' Abstract:
' mszip.vbs - compress a folder to a zipfile with directory hierarchy
'    or vice versa
'
' Usage:
' mszip {a | u} zipfile folder
'
' Examples:
' mszip a txt.zip dir1
' mszip u txt.zip dir2
'----------------------------------------------------------------------

WScript.Quit(VBSMain)

Sub ShowUsage()
	usage = "Compress a folder to a zipfile with directory hierarchy or vice versa" & vbCrlf & vbCrlf & _
		WScript.ScriptName & " {a | u} zipfile folder" & vbCrlf & _
		"  a" & Chr(9) & "compress folder to zipfile" & vbCrlf & _
		"  u" & Chr(9) & "uncompress zipfile to folder" & vbCrlf & _
		"  zipfile" & Chr(9) & "zipfile to compress to or to uncompress from" & vbCrlf & _
		"  folder" & Chr(9) & "target folder"
	WScript.Echo usage
End Sub

Function VBSMain()
	If CheckOSVersion = 0 Then
		WScript.Echo "Windows 2000 or later is required."
		VBSMain = 1 'Unsupported OS
		Exit Function
	End If

	argc = WScript.Arguments.Count
	If argc <> 3 Then
		VBSMain = 16 'unix man of zip's return value
		ShowUsage
		Exit Function
	End If

	Set argv = WScript.Arguments
	Select Case LCase(argv(0))
		Case "a"	sts = Zip(argv(1), argv(2))
		Case "u"	sts = UnZip(argv(1), argv(2))
		Case Else	sts = 16
	End Select
	VBSMain = sts
End Function

'Sub-Routines
Function CheckOSVersion()
On Error Resume Next
	os2k = "5.0"
	Set WSHShell = WScript.CreateObject("WScript.Shell") 
	strRegKey = "HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentVersion" 
	osVersion = WSHShell.RegRead(strRegKey)
	If Err.Number <> 0 Then
		Err.Clear
		CheckOSVersion = 0
		Exit Function
	End If
	If (osVersion >= os2k) Then CheckOSVersion = 1 Else CheckOSVersion = 0
End Function

Function NewBlankZip(pZipFile)
On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(pZipFile) Then fso.DeleteFile pZipFile, vbTrue
	Set blankzipfile = fso.OpenTextFile(pZipFile, 2, vbTrue)
	blankzipdata = "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
	blankzipfile.Write blankzipdata
	blankzipfile.Close
	Set fso = nothing :	Set blankzipfile = nothing
	NewBlankZip = 0
End Function

Function Zip(pZipFile, pSrcFolder)
On Error Resume Next
	zipfile = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(Replace(pZipFile, "/", "\"))
	srcfolder = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(Replace(pSrcFolder, "/", "\"))
	If NewBlankZip(zipfile) = 0 Then
		Set shlobj = CreateObject("Shell.Application")
		Set zipobj = shlobj.NameSpace(zipfile)
		Set srcobj = shlobj.NameSpace(srcfolder)
		'zipobj.CopyHere srcobj.Items
		skipped = 0
		For Each srcitem In srcobj.Items
			If srcitem.IsFolder Then
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set folderobj =	fso.GetFolder(srcitem.Path)
				If folderobj.Files.Count + folderobj.SubFolders.Count = 0 Then
					skipped = skipped + 1
				Else
					zipobj.CopyHere srcitem
				End If
			Else
				zipobj.CopyHere srcitem
			End If
		Next
		Do Until zipobj.Items.Count + skipped = srcobj.Items.Count
			WScript.Sleep 200
		Loop
		ret = 0
	Else
		ret = 1
	End If
	Zip = ret
End Function

Function CreateFolderTree(pDestFolder)
	Set fso = CreateObject("Scripting.FileSystemObject")
	parentfolder = fso.GetParentFolderName(pDestFolder)
	If Not fso.FolderExists(parentfolder) Then ret = CreateFolderTree(parentfolder)
	ret = fso.CreateFolder(pDestFolder)
	CreateFolderTree = ret
End Function

Function UnZip(pZipFile, pDestFolder)
On Error Resume Next
	zipfile = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(Replace(pZipFile, "/", "\"))
	destfolder = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(Replace(pDestFolder, "/", "\"))
	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(destfolder) Then CreateFolderTree destfolder
	Set shlobj = CreateObject("Shell.Application")
	Set zipitems = shlobj.NameSpace(zipfile).items
	UnZip = shlobj.NameSpace(destfolder).CopyHere(zipitems, 256)
End Function