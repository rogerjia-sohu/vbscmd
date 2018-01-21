With WScript .Quit(VBSMain(.Arguments.Count, .Arguments)) End With

Function VBSMain(argc, argv)
On Error Resume Next
	Dim strResult
	If argc = 2 Then
		Dim strData
		Dim bUseClipboard: bUseClipboard = False
		If  StrComp(argv(1), "/clip", vbTextCompare) = 0 Then
			strData = GetTextFromClipboard()
			If strData = vbNullString Then
				WScript.Echo "No text data available in clipboard! Quit!"
				VBSMain = 1
				Exit Function
			End If
			bUseClipboard = True
		Else
			strData = argv(1)
		End If

		Select Case argv(0)
		Case "e"
			strResult = base64Encode(strData)
		Case "d"
			strData = ConvertFromBase64EncodedUrl(strData)
			strResult = base64Decode(strData)
			strResult = ConvertToNormalUrl(strResult)
		Case Else
			strResult = vbNullString
			Err.Raise 5, argv(0), "Invalid command!"
		End Select
		If strResult <> vbNullString Then
			If bUseClipboard Then
				SetTextToClipboard(strResult)
			Else
				WScript.Echo strResult
			End If
			VBSMain = 0
		Else
			VBSMain = Err.Number
		End If
	Else
		VBSMain = base64Usage()
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

Function base64Usage()
	strMsg = "Usage: base64 cmd text" & vbNewLine _
		& vbTab & "cmd" & vbTab & "e to encode, d to decode.  One command at a time!" & vbNewLine  _
		& vbTab & "text" & vbTab & "Text. Use Clipboard I/O if it was specified to /clip ." & vbNewLine
	WScript.Echo strMsg
	base64Usage = 0
End Function

Function SetTextToClipboard(strData)
	Set objIE = CreateObject("InternetExplorer.Application") 
	objIE.Navigate("about:blank") 
	objIE.document.parentwindow.clipboardData.SetData "text", strData 
	objIE.Quit
	SetTextToClipboard = 1
End Function

Function GetTextFromClipboard()
	Dim strResult
	Set objIE = CreateObject("InternetExplorer.Application")
	objIE.Navigate("about:blank")
	strResult = objIE.document.parentwindow.clipboardData.GetData("text")
	objIE.Quit
	GetTextFromClipboard = strResult
End Function

Function ConvertFromBase64EncodedUrl(strBase64EncodedUrl)
	Dim url
	Dim thunderSig: thunderSig = "thunder://"
	Dim flashgetSig: flashgetSig = "Flashget://"

	If StrComp(Left(strBase64EncodedUrl, Len(thunderSig)), thunderSig, vbTextCompare) = 0 Then
		url = Replace(strBase64EncodedUrl, "thunder://", vbNullString, 1, -1, 1)
		If Right(url,1) = "/" Then url = Mid(url, 1, Len(url)-1)
	End If

	If StrComp(Left(strBase64EncodedUrl, Len(flashgetSig)), flashgetSig, vbTextCompare) = 0 Then
		url = Replace(strBase64EncodedUrl, "Flashget://", vbNullString, 1, -1, 1)
		If InStrRev(url,"&") > 0 Then url = Mid(url, 1, InStrRev(url,"&")-1)
	End If

	ConvertFromBase64EncodedUrl= url
End Function

Function ConvertFromFlashgetBase64Url(strFlashgetBase64Url)
	Dim url
	url = Replace(strFlashgetBase64Url, "Flashget://", vbNullString, 1, -1, 1)
	If InStrRev(url,"&") > 0 Then url = Mid(url, 1, InStrRev(url,"&")-1)
	ConvertFromFlashgetBase64Url = url
End Function

Function ConvertToNormalUrl(strBase64DecodedUrl)
	ConvertToNormalUrl = strBase64DecodedUrl
	If Len(strBase64DecodedUrl) > 4 And Left(strBase64DecodedUrl, 2) = "AA" And Right(strBase64DecodedUrl, 2) = "ZZ" Then
		ConvertToNormalUrl = Mid(strBase64DecodedUrl, 3, Len(strBase64DecodedUrl)-4)
		Exit Function
	End If
	If Len(strBase64DecodedUrl) > 20 And StrComp(Left(strBase64DecodedUrl,10), "[FLASHGET]" ,vbTextCompare) = 0 And _
		StrComp(Right(strBase64DecodedUrl,10), "[FLASHGET]", vbTextCompare) = 0 Then
		ConvertToNormalUrl = Mid(strBase64DecodedUrl, 11, Len(strBase64DecodedUrl)-20)
		Exit Function
	End If
End Function

Function Base64Encode(inData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut, sOut, I
  
  'For each group of 3 bytes
  For I = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _
      &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function

Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function

' Decodes a base-64 encoded string (BSTR type).
' 1999 - 2004 Antonin Foller, http://www.motobit.com
' 1.01 - solves problem with Access And 'Compare Database' (InStr)
Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string. (Length Error)"
    Base64Decode = vbNullString 'Added by jiatao 2010/07/10
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string. (" & thisChar & ")"
        Base64Decode = vbNullString 'Added by jiatao 2010/07/10
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function