VERSION 5.00
Begin VB.Form Customerlogin 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   9270
   ClientLeft      =   1125
   ClientTop       =   1305
   ClientWidth     =   17295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Customerlogin.frx":0000
   ScaleHeight     =   9270
   ScaleWidth      =   17295
   Begin VB.CommandButton Forgotpwbtn 
      BackColor       =   &H00C0C0FF&
      Caption         =   "FORGOT PASSWORD"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SHOW PASSWORD"
      Height          =   495
      Left            =   9240
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Exitbtn 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Submitbtn 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txtcus_id 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ID "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   2145
   End
End
Attribute VB_Name = "Customerlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
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

Private Sub Exitbtn_Click()
txtcus_id.Text = ""
txtpassword.Text = ""
If Check1.Value = 1 Then
    Check1.Value = 0
End If
Me.Hide
Startup.Show
End Sub


Private Sub Forgotpwbtn_Click()
Userforgot.Show
End Sub

Private Sub Submitbtn_Click()
rs1.Close
rs1.Open ("select * from customer where CUSTOMERID=" + txtcus_id + " And PASSWORD='" + txtpassword.Text + "' "), con, adOpenDynamic, adLockPessimistic

If Not rs1.EOF Then
    If Check1.Value = 1 Then
        Check1.Value = 0
    End If
    Me.Hide
    Home.Show
Else
    str = MsgBox("Invalid Customer Id or Passsword", vbExclamation + vbDefaultButton1)
End If

End Sub

Private Sub Form_Load()
con.Open ("provider=Microsoft.jet.OLEDB.4.0;data source=D:\Record.mdb;persist security info=false")
rs1.Open ("select * from customer"), con, adOpenDynamic, adLockPessimistic
End Sub


