VERSION 5.00
Begin VB.Form Userforgot 
   BackColor       =   &H00808000&
   Caption         =   "USER FORGOT FORM"
   ClientHeight    =   9165
   ClientLeft      =   300
   ClientTop       =   465
   ClientWidth     =   17565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   17565
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Caption         =   "FORGOT PASSWORD"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5655
      Left            =   6000
      TabIndex        =   0
      Top             =   1560
      Width           =   6375
      Begin VB.CommandButton Refreshbtn 
         BackColor       =   &H00C0E0FF&
         Caption         =   "REFRESH"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton Exitbtn 
         BackColor       =   &H00FFFFC0&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton Submitbtn 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtcus_id 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblpassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Password is:"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   3000
         Width           =   2010
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter CustomerID:"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Name:"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Top             =   1200
         Width           =   1440
      End
   End
End
Attribute VB_Name = "Userforgot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub Exitbtn_Click()
txtname.Text = ""
txtcus_id.Text = ""
lblpassword.Caption = ""
Me.Hide
End Sub

Private Sub Form_Load()
con.Open ("provider=Microsoft.jet.OLEDB.4.0;Data source=D:\Record.mdb;Persist security info=false")
rs.Open ("select * from customer "), con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Refreshbtn_Click()
txtname.Text = ""
txtcus_id.Text = ""
lblpassword.Caption = ""
End Sub

Private Sub Submitbtn_Click()
rs.Close
rs.Open ("select * from customer where NAME='" + txtname.Text + "'  and CUSTOMERID=" + txtcus_id.Text + " "), con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
    lblpassword.Caption = rs!Password
Else
    str = MsgBox("Details are Not Correct", vbExclamation)
End If
End Sub


