Const IMAPI_PROFILE_TYPE_CDROM = &H8
Const IMAPI_PROFILE_TYPE_DVDROM = &H10
Const IMAPI_PROFILE_TYPE_CD_RECORDABLE = &H9
Const IMAPI_PROFILE_TYPE_CD_REWRITABLE = &HA
Const IMAPI_PROFILE_TYPE_DVD_PLUS_RW = &H1A
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
objName = objFSO.GetTempName & ".htm"
objTempFile = objName

Set objFile = objFSO.CreateTextFile(objName, ForWriting, True)
Set colDiscMaster = CreateObject("IMAPI2.MsftDiscMaster2")

objFile.WriteLine "<html>"
objFile.WriteLine "<head>"
objFile.WriteLine "<title>CD Drive Information</title>"
objFile.WriteLine "</head><body>"
objFile.WriteLine "CD Drive Information:<p>"
For Each Id In colDiscMaster

    Set objRecorder = CreateObject("IMAPI2.MsftDiscRecorder2")
    objRecorder.InitializeDiscRecorder Id

'  objFile.WriteLine "--------------------------------<br>"
    objFile.WriteLine "Vendor: " & objRecorder.VendorId & "<br>"
    objFile.WriteLine "--------------------------------<br>"
    objFile.WriteLine "Product ID: " & objRecorder.ProductId & "<br>"
    objFile.WriteLine "--------------------------------<br>"
    objFile.WriteLine "Product Revision: " & objRecorder.ProductRevision & "<br>"
    For Each strMountPoint In objRecorder.VolumePathNames
      objFile.WriteLine "--------------------------------<br>"
        objFile.WriteLine "First Mount Point: " & strMountPoint & "<br>"
        Exit For
    Next

  objFile.WriteLine "--------------------------------<br>"
  objFile.WriteLine "Supported Profiles: <br>"
    For Each strProfile In objRecorder.SupportedProfiles
     Select Case strProfile
      Case IMAPI_PROFILE_TYPE_CDROM
        objFile.WriteLine "CD-ROM<br>"
     Case IMAPI_PROFILE_TYPE_DVDROM
         objFile.WriteLine "DVD-ROM<br>"
     Case IMAPI_PROFILE_TYPE_CD_RECORDABLE
         objFile.WriteLine "CD-R<br>"
        Case IMAPI_PROFILE_TYPE_CD_REWRITABLE
        objFile.WriteLine "CD-RW/CD+RW<br>"
     Case IMAPI_PROFILE_TYPE_DVD_PLUS_RW
        objFile.WriteLine "DVD+RW<br>"
   End Select
  Next
  objFile.WriteLine "--------------------------------<p>"
Next
'objFile.WriteLine "--------------------------------<br>"

objFile.WriteLine "</body></html>"
objFile.Close


Set objShell = CreateObject("Wscript.Shell")
objShell.Run objTempFile, 4, True
objFSO.DeleteFile(objTempFile)
