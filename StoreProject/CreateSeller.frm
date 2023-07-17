VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CreateSeller 
   Caption         =   "CreateSeller"
   ClientHeight    =   9975
   ClientLeft      =   525
   ClientTop       =   465
   ClientWidth     =   18885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "CreateSeller.frx":0000
   ScaleHeight     =   9975
   ScaleWidth      =   18885
   Begin VB.CommandButton Exit 
      BackColor       =   &H00CAC9EB&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Signupbtn 
      BackColor       =   &H00CAC9EB&
      Caption         =   "SIGNUP"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Loginbtn 
      BackColor       =   &H00CAC9EB&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CAC9EB&
      Caption         =   "NEW SELLER DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   7695
      Left            =   3120
      TabIndex        =   0
      Top             =   1320
      Width           =   11295
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9720
         Top             =   5640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtpan 
         Height          =   415
         Left            =   3240
         TabIndex        =   17
         Top             =   3720
         Width           =   2775
      End
      Begin VB.CommandButton Exitbtn 
         BackColor       =   &H0080C0FF&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6840
         Width           =   1935
      End
      Begin VB.CommandButton Submitbtn 
         BackColor       =   &H0080C0FF&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6840
         Width           =   1935
      End
      Begin VB.CommandButton UploadPhotobtn 
         BackColor       =   &H00C0FFFF&
         Caption         =   "UPLOAD PHOTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtemail 
         Height          =   415
         Left            =   3240
         TabIndex        =   12
         Top             =   6000
         Width           =   2775
      End
      Begin VB.TextBox txtmobno 
         Height          =   415
         Left            =   3240
         TabIndex        =   11
         Top             =   5280
         Width           =   2775
      End
      Begin VB.TextBox txtaadhar 
         Height          =   415
         Left            =   3240
         TabIndex        =   10
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtpassword 
         Height          =   415
         Left            =   3240
         TabIndex        =   9
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtaddress 
         Height          =   855
         Left            =   3240
         TabIndex        =   8
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtname 
         Height          =   415
         Left            =   3240
         TabIndex        =   7
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAN NUMBER:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   3720
         Width           =   1665
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   7680
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   6120
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   5280
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AADHAR NO:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   4560
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   960
         Width           =   780
      End
   End
End
Attribute VB_Name = "CreateSeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim str1 As String
Sub clear()
txtname.Text = ""
txtaddress.Text = ""
txtemail.Text = ""
txtpassword.Text = ""
txtpan.Text = ""
txtaadhar.Text = ""
txtmobno.Text = ""
Image1.Picture = LoadPicture("")
End Sub

Private Sub Exit_Click()
Me.Hide
Startup.Show
End Sub

Private Sub Loginbtn_Click()
SellerLogin.Show
Me.Hide
End Sub

Private Sub Signupbtn_Click()
Frame1.Visible = True

End Sub

Private Sub Submitbtn_Click()
If txtname.Text <> "" And txtaddress.Text <> "" And txtemail.Text <> "" And txtmobno.Text <> "" And txtpassword.Text <> "" And txtpan.Text <> "" And txtaadhar.Text <> "" And Image1.Picture <> LoadPicture("") Then
    rs.Close
    rs.Open ("select * from seller where EMAIL='" + txtemail.Text + "' And CONTACTNO='" + txtmobno + "' "), con, adOpenDynamic, adLockPessimistic
    If Not rs.EOF Then
        str = MsgBox("You already registered please login", vbInformation)
        clear
    Else
        rs.AddNew
        rs.Fields(1).Value = txtname.Text
        rs.Fields(2).Value = txtaddress.Text
        rs.Fields(3).Value = txtmobno.Text
        rs.Fields(4).Value = txtemail.Text
        rs.Fields(5).Value = txtaadhar.Text
        rs.Fields(6).Value = txtpan.Text
        rs.Fields(8).Value = txtpassword.Text
        rs.Fields(7).Value = str1
        str = MsgBox("Welcome '" & rs!Name & "' Your Id Number is '" & rs!id & "'", vbInformation + vbDefaultButton1)
        rs.Update
        clear
    End If
    ans = MsgBox("Do You want to login", vbQuestion + vbYesNo)
    If ans = vbYes Then
        Me.Hide
        SellerLogin.Show
    Else
        Me.Hide
    End If
Else
    str = MsgBox("Some Details Are Empty", vbExclamation)
End If
End Sub

Private Sub Exitbtn_Click()
Frame1.Visible = False
Me.Hide
clear
End Sub

Private Sub Uploadphotobtn_Click()
CommonDialog1.Filter = " jpg|*.jpg|jpeg|*.jpeg"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
str1 = CommonDialog1.FileName
Image1.Picture = LoadPicture(str1)
End Sub

Private Sub Form_Load()
con.Open ("provider=microsoft.jet.oledb.4.0;data source=D:\Record.mdb;persist security info=false")
rs.Open ("select * from SELLER"), con, adOpenDynamic, adLockPessimistic
Frame1.Visible = False
End Sub

