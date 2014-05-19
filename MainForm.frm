VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoSaver"
   ClientHeight    =   6060
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove program"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3180
      TabIndex        =   11
      Top             =   3900
      Width           =   1395
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Left            =   3420
      TabIndex        =   7
      Top             =   4380
      Width           =   1155
   End
   Begin VB.ListBox lstPrograms 
      Height          =   2985
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   1620
      TabIndex        =   9
      Top             =   5580
      Width           =   1215
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Auto-Save &Enabled"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   5100
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hi&de"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   10
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtApp 
      Height          =   345
      Left            =   900
      TabIndex        =   6
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox txtTime 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Text            =   "120"
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer tmrAutoSave 
      Interval        =   1000
      Left            =   240
      Top             =   5160
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add &new:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   4500
      Width           =   690
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "in these &programs:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1350
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "seconds"
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   180
      Width           =   585
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto &Save every"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuStartup 
         Caption         =   "Run on &Startup"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuLicence 
         Caption         =   "&Lincence agreement"
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMShow 
         Caption         =   "&Hide"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMEnabled 
         Caption         =   "&Enabled"
         Checked         =   -1  'True
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intTimer As Integer
Dim intListC As Integer

Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1

Private Sub cmdAdd_Click()
If Trim(Me.txtApp.Text) <> "" Then
    Me.lstPrograms.AddItem Trim(Me.txtApp.Text)
    Me.txtApp.Text = ""
    Me.cmdRemove.Enabled = True
End If
End Sub

Private Sub cmdHide_Click()
If frmAbout.Visible = True Then
    SetForegroundWindow frmAbout.hwnd
Else
    Me.Visible = Not Me.Visible
    If Me.Visible Then
        Me.mnuMShow.Caption = "&Hide"
    Else
        Me.mnuMShow.Caption = "&Show"
    End If
End If
End Sub

Private Sub cmdQuit_Click()
mnuQuit_Click
End Sub

Private Sub cmdRemove_Click()
If Me.lstPrograms.ListIndex < 0 Then
    Me.lstPrograms.ListIndex = Me.lstPrograms.ListCount - 1
End If

Me.lstPrograms.RemoveItem Me.lstPrograms.ListIndex

If Me.lstPrograms.ListCount <= 0 Then
    Me.cmdRemove.Enabled = False
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Integer

GetRegPos Me, False

Me.txtTime.Text = GetRegLong(HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "Time", Int(Me.txtTime.Text))
'Me.txtApp.Text = GetRegString(HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "Application", Me.txtApp.Text)

intListC = GetRegLong(HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "ProgramCount", 0)
For i = 0 To intListC - 1
    Me.lstPrograms.AddItem GetRegString(HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "Program" & i)
Next i

If intListC > 0 Then
    Me.cmdRemove.Enabled = True
End If

Set gSysTray = New clsSysTray
'Set gSysTray = New cTrayIcon
Set gSysTray.SourceWindow = Me
gSysTray.ToolTip = App.ProductName
gSysTray.ChangeIcon Me.Icon
gSysTray.DefaultDblClk = False

'gSysTray.Create Me.hwnd, Me.Icon.Handle, Me.Caption
gSysTray.IconInSysTray

'gSysTray.BalloonTipShow balIconInfo, "is running", Me.Caption, 5000

'Me.TrayIcon1.Create Me.hwnd, Me.Icon.Handle, Me.Caption
'Me.TrayIcon1.BalloonTipShow blIconInfo, "is running", "Crucigera Yelp!", 5000
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Cancel = True
    Me.Visible = False
    gSysTray.BalloonTipShow balIconInfo, "is still running", App.FileDescription
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim i As Integer

gSysTray.RemoveFromSysTray

SaveRegLong HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "Time", Val(Me.txtTime.Text)
'SaveRegString HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "Application", Me.txtApp.Text

SaveRegLong HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "ProgramCount", Me.lstPrograms.ListCount

For i = 0 To Me.lstPrograms.ListCount - 1
    SaveRegString HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.Title, "Program" & i, Me.lstPrograms.List(i)
Next i

SetRegPos Me, False
End Sub

Private Sub gSysTray_RButtonUp()
PopupMenu Me.mnuMenu, , , , Me.mnuMShow
End Sub

Private Sub chkEnabled_Click()
Me.tmrAutoSave.Enabled = Me.chkEnabled.Value
Me.mnuMEnabled.Checked = Me.tmrAutoSave.Enabled
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuLicence_Click()
LicenceForm.Show
End Sub

Private Sub mnuMAbout_Click()
mnuAbout_Click
End Sub

Private Sub mnuMEnabled_Click()
Me.tmrAutoSave.Enabled = Not Me.tmrAutoSave.Enabled
Me.chkEnabled.Value = -Me.tmrAutoSave.Enabled
Me.mnuMEnabled.Checked = Me.tmrAutoSave.Enabled
End Sub

Private Sub mnuMQuit_Click()
mnuQuit_Click
End Sub

Private Sub mnuMShow_Click()
cmdHide_Click
End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

Private Sub mnuStartup_Click()
MakeStartupReg "-silent"
End Sub

Private Sub tmrAutoSave_Timer()
Dim i As Integer
If intTimer >= Me.txtTime Then
    For i = 0 To Me.lstPrograms.ListCount - 1
        If InStr(1, GetActiveWindowTitle(False), Me.lstPrograms.List(i)) Then
            KeyDown vbKeyControl
            KeyDown vbKeyS
            KeyUp vbKeyS
            KeyUp vbKeyControl
            intTimer = intTimer Mod Me.txtTime + 1
        End If
    Next i
Else
    intTimer = intTimer + 1
End If
'Me.Caption = intTimer
End Sub

Private Sub gSysTray_LButtonDblClk()
cmdHide_Click
End Sub

Private Sub gSysTray_LButtonUp()
If Me.Visible = True Then SetForegroundWindow Me.hwnd
End Sub

Private Sub txtApp_Change()
If Trim(Me.txtApp.Text) <> "" Then
    Me.cmdAdd.Enabled = True
Else
    Me.cmdAdd.Enabled = False
End If
End Sub
