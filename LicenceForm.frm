VERSION 5.00
Begin VB.Form LicenceForm 
   Caption         =   "Licence Agreement"
   ClientHeight    =   5010
   ClientLeft      =   855
   ClientTop       =   1155
   ClientWidth     =   7935
   Icon            =   "LicenceForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox theContent 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "LicenceForm.frx":058A
      Top             =   600
      Width           =   7935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4935
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Also available at:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   2
      Top             =   210
      Width           =   1275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.justgroovy.net/groovy-general-free-software-licence"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   1455
      TabIndex        =   1
      Top             =   210
      Width           =   4755
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "LicenceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrH

GetRegPos Me
Me.Caption = "Licence Agreement - " & App.ProductName
Me.lblInfo(1).ToolTipText = "Go to " & Me.lblInfo(1).Caption

Exit Sub
ErrH:
errRes = MsgBox("Error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbApplicationModal + vbAbortRetryIgnore, App.FileDescription)
If errRes = vbRetry Then Resume
If errRes = vbIgnore Then Resume Next
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.theContent.Width = Me.ScaleWidth
Me.theContent.Height = Me.ScaleHeight - Me.theContent.Top
Me.cmdClose.Left = Me.ScaleWidth - Me.cmdClose.Width - 120
'Me.Line1.X2 = Me.ScaleWidth
'Me.Line2.X2 = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrH

SetRegPos Me

Exit Sub
ErrH:
errRes = MsgBox("Error #" & Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical + vbApplicationModal + vbAbortRetryIgnore, App.FileDescription)
If errRes = vbRetry Then Resume
If errRes = vbIgnore Then Resume Next
End Sub

Private Sub lblInfo_Click(Index As Integer)
If Left$(Me.lblInfo(Index).Caption, 7) = "http://" Then
    ShellExecute Me.hwnd, vbNullString, Me.lblInfo(Index).Caption, vbNullString, "C:\", SW_SHOWNORMAL
End If
End Sub

Private Sub theContent_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
    cmdClose_Click
End If
End Sub
