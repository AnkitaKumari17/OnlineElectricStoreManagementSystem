VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form Splash 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "SPLASH FORM"
   ClientHeight    =   9555
   ClientLeft      =   1350
   ClientTop       =   345
   ClientWidth     =   18630
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   477.75
   ScaleMode       =   0  'User
   ScaleWidth      =   931.5
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   15840
      Top             =   720
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   6120
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
      Max             =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ELECTRICAL STORE MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   585
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   9915
   End
   Begin VB.Label lblstat 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BorderColor     =   &H00404000&
      BorderWidth     =   3
      FillColor       =   &H00404000&
      Height          =   9495
      Left            =   120
      Top             =   0
      Width           =   18450
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
lblstatus.Caption = "Loading..Plz Wait.."
lblstat.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value = ProgressBar1.Max Then
    Timer1.Enabled = False
   ' Unload Me
    Startup.Show
End If
End Sub
