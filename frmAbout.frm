VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4740
   FillColor       =   &H80000005&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
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
      Height          =   435
      Left            =   3120
      TabIndex        =   0
      Top             =   5340
      Width           =   1515
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slogan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   6
      Left            =   135
      TabIndex        =   7
      Top             =   3060
      Width           =   645
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CC6600&
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   6
      Top             =   4500
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   120
      Picture         =   "frmAbout.frx":57E2
      Top             =   180
      Width           =   1920
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   4
      Left            =   135
      TabIndex        =   5
      Top             =   3420
      Width           =   4470
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   5000
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   0
      X2              =   5000
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   4
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   3
      Top             =   2790
      Width           =   630
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CC6600&
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2460
      Width           =   1110
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Groovy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   4995
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strComments() As String

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
Form_KeyPress KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
strComments = Split(App.Comments, ";")
Me.Caption = "About " & App.ProductName
Me.lblInfo(0).Caption = App.CompanyName
Me.lblInfo(1).Caption = App.Title
'Me.lblInfo(2).Left = Me.lblInfo(1).Left + Me.lblInfo(1).Width + 60
Me.lblInfo(2).Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & strComments(3)
Me.lblInfo(3).Caption = App.LegalCopyright
Me.lblInfo(4).Caption = Trim(strComments(1))
Me.lblInfo(5).Caption = Trim(strComments(2))
Me.lblInfo(5).ToolTipText = "Go to " & Me.lblInfo(5).Caption
Me.lblInfo(6).Caption = Trim(strComments(0))
End Sub

Private Sub lblInfo_Click(Index As Integer)
On Error Resume Next
If Left$(Me.lblInfo(Index).Caption, 7) = "http://" Then
    ShellExecute Me.hwnd, vbNullString, Me.lblInfo(Index).Caption, vbNullString, "C:\", SW_SHOWNORMAL
End If
End Sub
