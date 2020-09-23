Attribute VB_Name = "mCheck"
'--------------------------------------------------
' E-Mail Checker
' By H G Laughland
'
' Checking Module
' General declarations and procedures.
' Not all code is my own.
'--------------------------------------------------

Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_SYNC = &H0         '  play synchronously (default)

Public Type MailAccount
 tUserName As String
 tPassword As String
 tServer As String
End Type

Public ReturnCode As Long
Public blSwitchedOn As Boolean
Public blTimer As Boolean 'Holds last connection status.
Public blConnected As Boolean 'Holds connection status.
Public arServers() As MailAccount 'Holds the server list while program is active
Public iServerCount As Integer 'Holds current server total.
Public lInterval As Long 'Holds the interval between checks.

Public Sub ArrayLoad(strItem As String)
'Load data into an array element.
 iServerCount = iServerCount + 1
 ReDim Preserve arServers(iServerCount)
 arTemp = Split(strItem, "^", 3)
 arServers(iServerCount - 1).tServer = arTemp(0)
 arServers(iServerCount - 1).tUserName = arTemp(1)
 arServers(iServerCount - 1).tPassword = arTemp(2)
End Sub

Public Function ChooseFile(iType As Integer, cdTemp As CommonDialog) As String
'Common dialogue display.
Dim strFile As String
 cdTemp.Filter = "All Files (*.*)|*.*|TXT Files (*.txt)|*.txt"
 cdTemp.FilterIndex = 2
 Select Case iType
 Case 1 'Show openfile dialogue
  cdTemp.ShowOpen
 Case 2 'Show savefile dialogue
  cdTemp.ShowSave
 End Select
 ChooseFile = cdTemp.FileName
End Function

Public Function Connected_To_ISP() As Boolean
Dim hKey As Long
Dim lpSubKey As String
Dim phkResult As Long
Dim lpValueName As String
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As Long
Dim lpcbData As Long
 Connected_To_ISP = False
 lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
 ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, phkResult)
 If ReturnCode = ERROR_SUCCESS Then
  hKey = phkResult
  lpValueName = "Remote Connection"
  lpReserved = APINULL
  lpType = APINULL
  lpData = APINULL
  lpcbData = APINULL
  ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, _
  ByVal lpData, lpcbData)
  lpcbData = Len(lpData)
  ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, _
  lpData, lpcbData)
  If ReturnCode = ERROR_SUCCESS Then
   If lpData = 0 Then
    ' Not Connected
   Else
    ' Connected
    Connected_To_ISP = True
   End If
  End If
  RegCloseKey (hKey)
 End If
End Function

Function GetReg(hInKey As Long, ByVal subkey As String, ByVal valname As String)
  Dim RetVal As String, hSubKey As Long, dwType As Long
  Dim SZ As Long, v As String, r As Long
  RetVal = ""
  r = RegOpenKeyEx(hInKey, subkey, 0, 983139, hSubKey)
  If r <> 0 Then GoTo Ender
  SZ = 256: v = String(SZ, 0)
  r = RegQueryValueEx(hSubKey, valname, 0, dwType, ByVal v, SZ)
  If r = 0 And dwType = 1 Then
   RetVal = Left(v$, SZ - 1)
  Else
   RetVal = ""
  End If
  If hInKey = 0 Then r = RegCloseKey(hSubKey)
Ender:
  GetReg = RetVal
End Function

Private Function GetClient() As String
 Static strFolder As String, strKey As String
 strKey = "Software\Clients\Mail\"
 strResult = GetReg(&H80000002, strKey, "")
 strKey = strKey & strResult & "\Shell\Open\Command\"
 GetClient = GetReg(&H80000002, strKey, "")
End Function

Public Sub Main()
 Load frmMain
 lInterval = 300000
 If Not Dir(App.Path & "\Data.emc") = "" Then
  ReadData
  blSwitchedOn = True
  blConnected = Connected_To_ISP
  blTimer = True
  If blConnected Then
   frmMain.StartTimer
  Else
   frmMain.TimerAction2
  End If
  frmMain.ConnectionTimer
 Else
  frmSetUp.Show
 End If
End Sub

Public Sub ReadData()
'Read ata from the file.
 Dim strFile As String, strData As String
 Dim iHandle As Integer 'File handle
 strFile = App.Path & "\Data.emc"
 iHandle = FreeFile
 Open strFile For Input As #iHandle
 Do While Not EOF(iHandle)
  Line Input #iHandle, strData
  strData = BinaryToText(strData)
  ArrayLoad strData
 Loop
 Close #iHandle
End Sub

Public Sub RunMail()
Dim strTemp As String
 strTemp = GetClient
 Shell strTemp, vbNormalFocus
End Sub

Public Sub SaveFile()
'Save Data To File
Dim F As Integer, x As Integer
Dim response As Integer, strTemp As String, strFile As String
On Error GoTo CloseError
If Not iServerCount > 0 Then Exit Sub
 strFile = App.Path & "\Data.emc"
 F = FreeFile
 Open strFile For Output As F
 For x = 0 To iServerCount - 1
  strTemp = arServers(x).tServer & "^" & arServers(x).tUserName & "^" & arServers(x).tPassword
  strTemp = TextToBinary(strTemp)
  Print #F, strTemp
 Next
 Close F
 Exit Sub
CloseError:
 MsgBox "Error occurred while trying to close file, please retry.", 48
End Sub

Public Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
 frmMain.TimerAction
End Sub

Public Sub TimerProc2(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
 If blSwitchedOn Then
  blConnected = Connected_To_ISP
  frmMain.TimerAction2
 End If
End Sub

Public Sub TimerProc3(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
 frmMain.TimerAction3
End Sub

