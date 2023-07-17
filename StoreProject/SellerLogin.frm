VERSION 5.00
Begin VB.Form SellerLogin 
   BackColor       =   &H00808000&
   Caption         =   "SELLER LOGIN FORM"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "SellerLogin.frx":0000
   ScaleHeight     =   8940
   ScaleWidth      =   17700
   Begin VB.Frame Frame1 
      BackColor       =   &H00CAC9EB&
      Caption         =   "SELLER LOGIN"
      Height          =   5295
      Left            =   4080
      TabIndex        =   0
      Top             =   1440
      Width           =   7815
      Begin VB.CheckBox Check1 
         BackColor       =   &H00D9FEFF&
         Caption         =   "Show password"
         Height          =   435
         Left            =   4080
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         MaskColor       =   &H00DDFEFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   600
         MaskColor       =   &H00DDFEFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtpassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtid 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "FORGOT PASSWORD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   1080
         TabIndex        =   6
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIN ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "SellerLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
 
Private Sub Check1_Click()
If Check1.Value = 1 Then
    txtpassword.PasswordChar = ""
    Check1.Caption = "HIDE PASSWORD"
Else
    txtpassword.PasswordChar = "*"
    Check1.Caption = "SHOW PASSWORD"
End If
End Sub
 
Private Sub Command1_Click()
rs.Close
rs.Open ("select * from seller where ID=" + txtid.Text + " and PASSWORD='" + txtpassword.Text + "' "), con, adOpenDynamic, adLockPessimistic
   If Not rs.EOF Then
        'txtid.Text = ""
        'txtpassword.Text = ""
        If Check1.Value = 1 Then
            Check1.Value = 0
        End If
        Me.Hide
        SellerHome.Show
    
    Else
    str = MsgBox("Invalid Seller ID Or PASSWORD", vbExclamation)
    End If
   
End Sub
Private Sub Command2_Click()
txtid.Text = ""
txtpassword.Text = ""
Check1.Value = 0
Frame1.Visible = False
Startup.Show
Me.Hide
End Sub

Private Sub Form_Load()
Frame1.Visible = True
con.Open ("provider=microsoft.jet.oledb.4.0;data source=D:\Record.mdb ;persist security info=false")
rs.Open ("select * from Seller"), con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Label3_Click()
OwnerForgot.Show
End Sub



