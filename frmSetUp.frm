VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-Mail Checker Setup"
   ClientHeight    =   3375
   ClientLeft      =   3165
   ClientTop       =   2940
   ClientWidth     =   4950
   Icon            =   "frmSetUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   3720
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSetUp.frx":0442
      Left            =   2160
      List            =   "frmSetUp.frx":045B
      TabIndex        =   10
      Text            =   "5"
      Top             =   2760
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox cmbServer 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   120
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Label5 
      Caption         =   "Check every"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   2800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "minutes"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   2800
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1100
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Username:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   740
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "frmSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------
' E-Mail Checker
' By H G Laughland
'
' Settings Form
'
' This form is used to configure the checker.
' Some code is not original. My thanks to the
' authors concerned.
'--------------------------------------------------

Private Sub cmbServer_Click()
Dim strTemp As String, iTemp As Integer
 iTemp = FindRecord
 Me.cmbServer.Text = arServers(iTemp).tServer
 Me.txtUser.Text = arServers(iTemp).tUserName
 Me.txtPass.Text = arServers(iTemp).tPassword
End Sub

Private Sub cmdAdd_Click()
 Me.cmbServer.Text = ""
 Me.txtUser.Text = ""
 Me.txtPass.Text = ""
End Sub

Private Sub cmdClose_Click()
 
 Unload Me
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you sure you want to delete this record?", 4 + 32, "Delete Record") = vbYes Then DeleteItem
End Sub

Private Sub cmdUpdate_Click()
Dim strTemp As String, iTemp As Integer
iTemp = FindRecord
If iTemp < 0 Then
 strTemp = Me.cmbServer.Text & "^" & Me.txtUser.Text & "^" & Me.txtPass.Text
 ArrayLoad strTemp
 Me.cmbServer.AddItem Me.cmbServer.Text
Else
 UpdateItem iTemp
End If
End Sub

Private Sub Combo1_Click()
 lInterval = CLng(Combo1.List(Combo1.ListIndex)) * 1000
End Sub

Private Sub Form_Load()
Dim x As Integer
 KillTimer frmMain.hwnd, 2
 If iServerCount > 0 Then
  For x = 0 To iServerCount - 1
   cmbServer.AddItem arServers(x).tServer
  Next
  cmbServer.Text = arServers(0).tServer
  txtUser.Text = arServers(0).tUserName
  txtPass.Text = arServers(0).tPassword
 End If
End Sub

Private Sub Form_Paint()
 DrawFrameOn Image1, Image1, "outward", 4
 DrawFrameOn Image1, Image1, "inward", 1
 DrawFrameOn Image2, Image2, "outward", 4
 DrawFrameOn Image2, Image2, "inward", 1
End Sub

Private Sub DrawFrameOn(TopLeftControl As Control, LowestRightControl As Control, Style As String, Framewidth As Integer)
Dim dw, fs, sm
Dim st$
Dim Lft, Toplft, Hite
Dim Rite, Ritebotm
Dim lt As Long
Dim rb As Long

    'Routine to draw frame around controls.  Variations can be achieved by changing "DrawWidth", shadow colors
    ' and the width of the frame elements in the form Paint Event
    
    'Save the current settings
    dw = DrawWidth
    fs = FillStyle
    sm = ScaleMode
    
    DrawWidth = 1
    FillStyle = 1
    ScaleMode = 3
    
    st = LCase(Left$(Style, 1))
    Lft = TopLeftControl.Left
    Toplft = TopLeftControl.Top
    Hite = TopLeftControl.Height
    
    Rite = LowestRightControl.Left + LowestRightControl.Width
    Ritebotm = LowestRightControl.Top + LowestRightControl.Height
    
    If Ritebotm > Hite Then Hite = Ritebotm
       
    lt = vb3DHighlight
    rb = vbButtonShadow
    
    'Swap colors if "inward"
    If st = "i" Then
        lt = vb3DDKShadow
        rb = vb3DHighlight
    End If
    
    'Draw the frame
    Line (Lft - Framewidth, Toplft - Framewidth)-(Rite + Framewidth, Toplft - Framewidth), lt
    Line (Lft - Framewidth, Toplft - Framewidth)-(Lft - Framewidth, Hite + Framewidth), lt
    Line (Rite + Framewidth, Toplft - Framewidth)-(Rite + Framewidth, Ritebotm + Framewidth), rb
    Line (Rite + Framewidth, Ritebotm + Framewidth)-(Lft - Framewidth, Hite + Framewidth), rb
    
    'Restore original settings
    DrawWidth = dw
    FillStyle = fs
    
    ScaleMode = sm
     
End Sub

Private Sub UpdateItem(iEntry As Integer)
 arServers(iServerCount - 1).tServer = cmbServer.Text
 arServers(iServerCount - 1).tUserName = txtUser.Text
 arServers(iServerCount - 1).tPassword = txtPass.Text
End Sub

Private Function FindRecord() As Integer
Dim x As Integer, iTemp As Integer
iTemp = -1
For x = 0 To iServerCount - 1
 If arServers(x).tServer = Me.cmbServer.Text Then iTemp = x
Next
 FindRecord = iTemp
End Function

Public Sub DeleteItem()
'Delete Selected Record
Dim iTemp As Integer, iTemp2 As Integer, x As Integer
iTemp = FindRecord
iTemp2 = 0
While Not cmbServer.List(iTemp2) = cmbServer.Text
 iTemp2 = iTemp2 + 1
Wend
For x = iTemp To iServerCount - 1
 If x < iServerCount - 1 Then
  arServers(x).tServer = arServers(x + 1).tServer
  arServers(x).tUserName = arServers(x + 1).tUserName
  arServers(x).tPassword = arServers(x + 1).tPassword
 End If
Next
iServerCount = iServerCount - 1
ReDim Preserve arServers(iServerCount)
x = iTemp2
cmbServer.RemoveItem x
If x = 0 Then x = 1
Me.cmbServer.Text = arServers(x - 1).tServer
Me.txtUser.Text = arServers(x - 1).tUserName
Me.txtPass.Text = arServers(x - 1).tPassword
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmMain.StartTimer
End Sub
