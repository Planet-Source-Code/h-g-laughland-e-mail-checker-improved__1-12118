VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-Mail Checker"
   ClientHeight    =   3030
   ClientLeft      =   1725
   ClientTop       =   2940
   ClientWidth     =   8925
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Read Mail"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   8435
      _ExtentX        =   14870
      _ExtentY        =   3836
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Server"
         Object.Width           =   3233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "#"
         Object.Width           =   1028
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   3233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Address"
         Object.Width           =   2792
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Subject"
         Object.Width           =   4380
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   2520
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSwitch 
         Caption         =   "&Pause"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "&Check Now"
      End
      Begin VB.Menu mnuMailList 
         Caption         =   "&Mail List"
      End
      Begin VB.Menu mnuRead 
         Caption         =   "&Read Mail"
      End
      Begin VB.Menu mnuSetUp 
         Caption         =   "Set&Up"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------
' E-Mail Checker
' By H G Laughland
'
' Main Form
'
' This form only appears as a system tray icon.
' Some code is not original. My thanks to the
' authors concerned.
'--------------------------------------------------
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const LVM_FIRST = &H1000

Private Enum POP3States
 POP3_Connect
 POP3_USER
 POP3_PASS
 POP3_STAT
 POP3_RETR
 POP3_DELE
 POP3_QUIT
 POP3_TOP
End Enum

Private Type RecdMsg
 strServer As String
 iItem As Integer
End Type

Dim rmMsg() As RecdMsg
Private m_State As POP3States
Dim strServer As String
Dim sFrom As String, sSubject As String
Dim pMsg() As String, iMailCount As Integer
Dim iCheckServer As Integer, iMailRecd As Integer, strFrom As String
Dim li As ListItem, blSelected As Boolean
Dim intMessages As Integer 'the number of messages to be loaded

Private Sub cmdCheck_Click()
 Me.Hide
 mnuRead_Click
End Sub

Private Sub Command1_Click()
 Me.Hide
End Sub

Private Sub Form_Load()
 Me.Hide
 SetTrayIcon Me, "E-Mail Checker", LoadResPicture(101, 1), 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Message As Long
On Error Resume Next
 Message = x / Screen.TwipsPerPixelX
 Select Case Message
  Case WM_RBUTTONUP
   PopupMenu mnuPopUp
  Case WM_RBUTTONDOWN
   PopupMenu mnuPopUp
  Case WM_LBUTTONDBLCLK
   Me.Show
 End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 SaveFile
 RemoveTrayIcon
End Sub

Private Sub mnuCheck_Click()
 TimerAction
End Sub

Private Sub mnuMailList_Click()
 Me.Show
End Sub

Private Sub mnuQuit_Click()
 SaveFile
 RemoveTrayIcon
 End
End Sub

Private Sub mnuRead_Click()
 SetTrayIcon Me, "E-Mail Checker", LoadResPicture(101, 1), 2
 RunMail
End Sub

Private Sub mnuSetUp_Click()
 frmSetUp.Show
End Sub

Private Sub mnuSwitch_Click()
Select Case blSwitchedOn
 Case True
  mnuSwitch.Checked = True
  blSwitchedOn = False
  blConnected = False
  TimerAction2
 Case False
  mnuSwitch.Checked = False
  blSwitchedOn = True
End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String

'Save the received data into strData variable
strData = ""
Winsock1.GetData strData

 Select Case m_State
  
  Case POP3_Connect, POP3_USER, POP3_PASS
   Login strData
            
  Case POP3_STAT, POP3_TOP
   CheckMessages strData
  
  Case POP3_QUIT
   'Close the connection
   Debug.Print "BYE"
   Winsock1.Close
   If Not blSelected Then CheckMail
 End Select
 
End Sub

Private Sub CommenceCheck(iServers As Integer)
  'POP3 server waits for the connection request at the port 110.
  SetTrayIcon Me, "Checking " & arServers(iCheckServer).tServer, LoadResPicture(102, 1), 2
  strServer = arServers(iCheckServer).tServer
  ConnectToServer strServer
End Sub

Private Sub CheckMail()
' Check all pop servers.
iCheckServer = iCheckServer + 1
If iCheckServer < iServerCount Then
 CommenceCheck iCheckServer
Else
 If iMailRecd > 0 Then
  SetTrayIcon Me, "You have " & iMailRecd & " messages waiting", LoadResPicture(103, 1), 2
  sndPlaySound App.Path & "\Notify.wav", SND_FILENAME Or SND_SYNC
 Else
  SetTrayIcon Me, "E-Mail Checker", LoadResPicture(101, 1), 2
 End If
End If
End Sub

Public Sub CheckMessages(strData As String)
Dim iPos As Integer, Reply As String
Static iHeader As Integer 'Header counter

'Check if mail waiting and process headers.
Select Case m_State

 Case POP3_STAT
   'Check what is returned by STAT to see if
   'there are messages waiting.
   intMessages = CInt(Mid$(strData, 5, _
   InStr(5, strData, " ") - 5))
   If intMessages > 0 Then
    'There is something in the mailbox!
    iMailRecd = iMailRecd + intMessages
    iHeader = 1
    m_State = POP3_TOP
    Winsock1.SendData "TOP " & iHeader & " 0" & vbCrLf
   Else
    'The mailbox is empty.
    m_State = POP3_QUIT
    Winsock1.SendData "QUIT" & vbCrLf
    Debug.Print "QUIT"
   End If
  
 Case POP3_TOP
   'Debug.Print strData
   Reply = strData
   iPos = InStr(1, UCase(Reply), "FROM:")
   If iPos Then
    ExtractHeader Reply, strServer, iHeader
    iHeader = iHeader + 1
    If iHeader > intMessages Then
     m_State = POP3_QUIT
     Winsock1.SendData "QUIT " & vbCrLf
    Else
     Winsock1.SendData "TOP " & iHeader & " 0" & vbCrLf
    End If
   End If
  
End Select
End Sub

Public Sub ColumnResize()
Dim x As Integer
For x = 1 To 5
 Call SendMessage(ListView1.hwnd, LVM_FIRST + 30, x - 1, -2)
Next
End Sub

Public Sub ConnectionTimer()
 SetTimer Me.hwnd, 1, 1000, AddressOf TimerProc2
End Sub



Public Sub ExtractHeader(Reply As String, Server As String, iItem As Integer)
Dim iPos As Integer, iPos2 As Integer
Dim strName As String, strAddress As String

 sFrom = "": sSubject = ""
 iPos = InStr(1, UCase(Reply), "FROM:")
 
 'Extract sender and subject details
 If iPos Then
  sFrom = Mid(Reply, iPos + 6)
  sFrom = Left(sFrom, InStr(1, sFrom, vbCrLf) - 1)
  If Left(sFrom, 1) = vbLf Then sFrom = ""
  Debug.Print "From: " & sFrom
  iPos = InStr(1, UCase(Reply), "SUBJECT:")
  If iPos Then
   sSubject = Mid(Reply, iPos + 9)
   sSubject = Left(sSubject, InStr(1, sSubject, vbCrLf) - 1)
   If Left(sSubject, 1) = vbLf Then sSubject = ""
   Debug.Print "Subject: " & sSubject
  End If
 End If
 
 ' Load the mail list screen
 If sFrom <> "" Then
 ' Separate sender into name & address
  iPos = InStr(1, sFrom, "<")
  iPos2 = InStr(1, sFrom, ">")
  If iPos > 0 Then
   strName = Mid(sFrom, 1, iPos - 1)
   strAddress = Mid(sFrom, iPos + 1, iPos2 - iPos - 1)
  Else
   iPos = InStr(1, sFrom, " ")
   strName = Mid(sFrom, 1, iPos)
   strAddress = Mid(sFrom, 1, Len(sFrom) - iPos)
  End If
  ' Add to listview
  Set li = ListView1.ListItems.Add(, , strServer)
  li.ListSubItems.Add , , iItem
  li.ListSubItems.Add , , strName
  li.ListSubItems.Add , , strAddress
  li.ListSubItems.Add , , sSubject
 End If
 ColumnResize
End Sub

Public Sub Login(strData As String)
'Handles logging in to server
Select Case m_State
 
 Case POP3_Connect
   'Reset the number of messages
   intMessages = 0
   m_State = POP3_USER
   'Send to the server the USER command with the parameter.
   Winsock1.SendData "USER " & arServers(iCheckServer).tUserName & vbCrLf
   
 Case POP3_USER
   m_State = POP3_PASS
   Winsock1.SendData "PASS " & arServers(iCheckServer).tPassword & vbCrLf
   
 Case POP3_PASS
  m_State = POP3_STAT
   'Send STAT command to know how many messages in the mailbox
   Winsock1.SendData "STAT" & vbCrLf

 End Select

End Sub

Public Sub StartTimer()
'Start periodic mail checking.
 SetTimer Me.hwnd, 2, lInterval, AddressOf TimerProc
 TimerAction
End Sub

Public Sub TimerAction()
 If blConnected Then
  SetTrayIcon Me, "Checking...", LoadResPicture(102, 1), 2
  iCheckServer = -1
  iMailRecd = 0
  ListView1.ListItems.Clear
  CheckMail
 End If
End Sub

Public Sub TimerAction2()
If blConnected Then
 If Not blTimer Then
  blTimer = blConnected
  SetTrayIcon Me, "E-Mail Checker", LoadResPicture(101, 1), 2
  CheckMail
  StartTimer
 End If
Else
 If blTimer Then
  blTimer = blConnected
  SetTrayIcon Me, "Disabled", LoadResPicture(104, 1), 2
  KillTimer Me.hwnd, 2
 End If
End If
End Sub

Public Sub TimerAction3()
 KillTimer Me.hwnd, 3
 If m_State = POP3_Connect Then CheckMail
End Sub

Public Sub ConnectToServer(strServer As String)
'Change the value of current session state
   m_State = POP3_Connect
    
  'Close the socket in case it was opened while another session
  Winsock1.Close
    
  'Reset the value of the local port.
  Winsock1.LocalPort = 0
  
  Winsock1.Connect strServer, 110
  SetTimer Me.hwnd, 3, 30000, AddressOf TimerProc3
End Sub

