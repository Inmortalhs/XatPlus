Attribute VB_Name = "Source"
Option Explicit
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
       
Private Const BUFFER_LEN = 256
Public Function GetUrlSource(sURL As String) As String

    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long

    'get the handle of the current internet connection
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    'get the handle of the url
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    'if we have the handle, then start reading the web page
    If hInternet Then
        'get the first chunk & buffer it.
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        'if there's more data then keep reading it into the buffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    'close the URL
    iResult = InternetCloseHandle(hInternet)

    GetUrlSource = sData

End Function
Public Function StrFilter(ByVal Text As String, ByVal Chars As String, Optional ByVal PassThru As Boolean = False, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As String
Dim i As Long

If PassThru Then
For i = 1 To Len(Text)
If InStr(1, Chars, Mid$(Text, i, 1), Compare) Then
StrFilter = StrFilter & Mid$(Text, i, 1)
End If
Next i
Else
For i = 1 To Len(Text)
If InStr(1, Chars, Mid$(Text, i, 1), Compare) = 0 Then
StrFilter = StrFilter & Mid$(Text, i, 1)
End If
Next i
End If
End Function
Public Function VL64Encode(ByVal data As Long) As String
On Error GoTo efun

Dim wf(6) As Byte, pos As Integer, startpos As Integer, bytes As Integer, negmask As Integer
pos = 0
startpos = pos
bytes = 1
If data >= 0 Then negmask = 0 Else negmask = 4

data = Math.Abs(data)
wf(pos) = 64 + (data Mod 4)
pos = pos + 1

data = Math.Round((data / 4) - 0.49)
While data <> 0
bytes = bytes + 1
wf(pos) = 64 + (data Mod 64)
pos = pos + 1
data = Math.Round((data / 64) - 0.49)
Wend
wf(startpos) = wf(startpos) Or bytes * 8 Or negmask

Dim tmp As String, j As Integer
tmp = vbNullString

For j = 0 To 6
tmp = tmp & Chr$(wf(j))
Next

If InStr(tmp, Chr(0)) Then
tmp = StrFilter(tmp, Chr(0), False)
End If

VL64Encode = tmp
Exit Function

efun:
VL64Encode = vbNullString

End Function
Public Function VL64Decode(data As String) As String
On Error Resume Next

Dim Second As String
Dim nf As Long
Dim i As Long
Dim R As Long
Dim ID As String
Dim rawid As String

nf = (Asc(Left(data, 1)) - 72) / 4
If nf < 1 Then
nf = 0
End If
If (nf Mod 2) = 0 Then
nf = nf / 2
For i = 1 To nf
R = R + ((Asc(Mid(data, i + 1, 1)) - 64) * (64 ^ (i - 1)))
Next i

ID = (4 * R) + ((Asc(Left(data, 1)) - 72) Mod 4)
Else
rawid = Left(data, 2)
Second = Replace(rawid, "S", "", 1, 1)
ID = (Asc(Second) - 64) * 4 + nf
End If
VL64Decode = ID
End Function
Public Sub Pause(Duration)
  Dim numTime
  numTime = Timer
  Do While Timer - numTime < Duration
    DoEvents
  Loop
End Sub



