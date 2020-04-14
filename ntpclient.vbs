Option Explicit

Dim sUrl
sUrl = "http://ntp-a1.nict.go.jp/cgi-bin/ntp"

Dim oWinHttpReq 
Set oWinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")

oWinHttpReq.Open "GET", sUrl, false
oWinHttpReq.Send

Dim arrRes
arrRes = Split(Replace(Replace(oWinHttpReq.ResponseText, "<", ","), ">", ","), ",")

Dim v, raw
For Each v In arrRes
	If IsNumeric(v) Then raw = v
Next

Dim NewDate
NewDate = DateAdd("n", CDbl(raw/60), CDate("1900-01-01 00:00:00"))

Dim fv
fv = raw / 1000

Dim s
s = fix(1000 * (fv - Fix(fv))) Mod 60
NewDate = DateAdd("h", 9, DateAdd("s", s, NewDate))

Dim arrDt
arrDt = Split(NewDate)

Dim WshShell 
Set WshShell = CreateObject("wscript.Shell") 
WshShell.Run "cmd.exe /c date " & arrDt(0), 0 
WshShell.Run "cmd.exe /c time " & arrDt(1), 0 
