VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CreateCustomer 
   Caption         =   "CUSTOMER DETAIL FORM"
   ClientHeight    =   9945
   ClientLeft      =   825
   ClientTop       =   1005
   ClientWidth     =   17700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "CreateCustomer.frx":0000
   ScaleHeight     =   9945
   ScaleWidth      =   17700
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12840
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   8535
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin VB.TextBox txtconfirmpw 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   27
         Top             =   7080
         Width           =   2775
      End
      Begin VB.CommandButton Customerbtn 
         BackColor       =   &H0080FFFF&
         Caption         =   "ALREADY CUSTOMER"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   7680
         Width           =   1935
      End
      Begin VB.CommandButton Uploadpanbtn 
         BackColor       =   &H00FFFFC0&
         Caption         =   "UPLOAD PAN"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   475
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton Uploadaadharbtn 
         BackColor       =   &H00FFFFC0&
         Caption         =   "UPLOAD AADHAR"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   475
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5280
         Width           =   2175
      End
      Begin VB.CommandButton Cancelbtn 
         BackColor       =   &H006A6FC1&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   7680
         Width           =   1935
      End
      Begin VB.CommandButton Submitbtn 
         BackColor       =   &H0080FFFF&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   7680
         Width           =   2055
      End
      Begin VB.CommandButton Uploadphotobtn 
         BackColor       =   &H0080FFFF&
         Caption         =   "UPLOAD PHOTO"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtpassword 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   20
         Top             =   6480
         Width           =   2775
      End
      Begin VB.TextBox txtemail 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   19
         Top             =   5880
         Width           =   2775
      End
      Begin VB.TextBox txtaadharno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   18
         Top             =   5280
         Width           =   2775
      End
      Begin VB.TextBox txtpanno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   17
         Top             =   4680
         Width           =   2775
      End
      Begin VB.TextBox txtphone 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   16
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txtpincode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   15
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox txtstate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   14
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtdistrict 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   13
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtaddress 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2880
         TabIndex        =   12
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIRM PASSWORD:"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   28
         Top             =   7080
         Width           =   2340
      End
      Begin VB.Image Panimg 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Image Aadharimg 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Image Photoimg 
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   5880
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AADHAR NO:"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   5280
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAN NUMBER :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO. :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   4080
         Width           =   1545
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PIN CODE :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   3480
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STATE :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   2880
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DISTRICT :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   765
      End
   End
End
Attribute VB_Name = "CreateCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim str1 As String
Dim str2 As String
Dim str3 As String

Private Sub Cancelbtn_Click()
    txtname.Text = ""
    txtaddress.Text = ""
    txtdistrict.Text = ""
    txtstate.Text = ""
    txtpincode.Text = ""
    txtphone.Text = ""
    txtpanno.Text = ""
    txtaadharno.Text = ""
    txtemail.Text = ""
    txtpassword.Text = ""
    txtconfirmpw.Text = ""
    Photoimg.Picture = LoadPicture("")
    Panimg.Picture = LoadPicture("")
    Aadharimg.Picture = LoadPicture("")
    Me.Hide
End Sub

Private Sub Customerbtn_Click()
Me.Hide
Customerlogin.Show
End Sub

Private Sub Form_Load()
con.Open ("provider=Microsoft.jet.OLEDB.4.0; Data Source=D:\Record.mdb; Persist security info=false")
rs.Open ("select * from customer"), con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Submitbtn_Click()
If txtname.Text <> "" And txtaddress.Text <> "" And txtdistrict.Text <> "" And txtstate.Text <> "" And txtpincode.Text <> "" And txtphone.Text <> "" And txtpanno.Text <> "" And txtaadharno.Text <> "" And txtemail.Text <> "" And txtpassword.Text <> "" And txtconfirmpw.Text <> "" Then
    If Photoimg.Picture <> LoadPicture("") And Panimg.Picture <> LoadPicture("") And Aadharimg.Picture <> LoadPicture("") Then
        If txtpassword.Text <> txtconfirmpw.Text Then
            str = MsgBox("Password missmatched", vbExclamation)
            txtconfirmpw.Text = ""
            str = MsgBox("Re-enter password", vbInformation)
            Me.txtconfirmpw.SetFocus
            Exit Sub
        Else
        
        rs.AddNew
        rs.Fields(1).Value = txtname.Text
        rs.Fields(2).Value = txtaddress.Text
        rs.Fields(3).Value = txtdistrict.Text
        rs.Fields(4).Value = txtstate.Text
        rs.Fields(5).Value = txtpincode.Text
        rs.Fields(6).Value = txtphone.Text
        rs.Fields(7).Value = txtpanno.Text
        rs.Fields(8).Value = txtaadharno.Text
        rs.Fields(9).Value = txtemail.Text
        rs.Fields(10).Value = txtpassword.Text
        rs.Fields(14).Value = txtconfirmpw.Text
        rs.Fields(11).Value = str1
        rs.Fields(12).Value = str2
        rs.Fields(13).Value = str3
        str = MsgBox("Welcome" & rs!Name & " Your Customer id is '" & rs!Customerid & "'Now you login from Already Customer Button.Thankyou  ", vbInformation + vbDefaultButton1)
        rs.Update
        
        txtname.Text = ""
        txtaddress.Text = ""
        txtdistrict.Text = ""
        txtstate.Text = ""
        txtpincode.Text = ""
        txtphone.Text = ""
        txtpanno.Text = ""
        txtaadharno.Text = ""
        txtemail.Text = ""
        txtpassword.Text = ""
        txtconfirmpw.Text = ""
        Photoimg.Picture = LoadPicture("")
        Panimg.Picture = LoadPicture("")
        Aadharimg.Picture = LoadPicture("")
        End If
        
    Else
        str = MsgBox("Document or Photos Not Uploaded ", vbExclamation + vbOKOnly)
    End If
Else
    str = MsgBox("some Columns are empty ", vbExclamation)
End If

End Sub



Private Sub Uploadaadharbtn_Click()
CommonDialog1.Filter = "jpeg|*.jpg"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
str3 = CommonDialog1.FileName
Aadharimg.Picture = LoadPicture(str3)
End Sub

Private Sub Uploadpanbtn_Click()
CommonDialog1.Filter = "jpg|*.jpg|jpeg|*.jpeg"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen
str2 = CommonDialog1.FileName
Panimg.Picture = LoadPicture(str2)
End Sub

Private Sub Uploadphotobtn_Click()
CommonDialog1.Filter = "jpeg|*.jpg"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
str1 = CommonDialog1.FileName
Photoimg.Picture = LoadPicture(str1)
End Sub
