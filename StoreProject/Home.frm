VERSION 5.00
Begin VB.Form Home 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9900
   ClientLeft      =   120
   ClientTop       =   570
   ClientWidth     =   18960
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   18960
   Begin VB.CommandButton Cartbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "CART"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   7
      Left            =   13560
      TabIndex        =   31
      Top             =   5280
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   7
         Left            =   840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblpname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label lblprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   33
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblpdetail 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   32
         Top             =   3960
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   6
      Left            =   9240
      TabIndex        =   27
      Top             =   5280
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   6
         Left            =   840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblpname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   30
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label lblprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   29
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblpdetail 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   3960
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   5
      Left            =   4920
      TabIndex        =   23
      Top             =   5280
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   5
         Left            =   840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblpname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   26
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label lblprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   5
         Left            =   600
         TabIndex        =   25
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblpdetail 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   24
         Top             =   3960
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   4
      Left            =   600
      TabIndex        =   19
      Top             =   5280
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   4
         Left            =   840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblpname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label lblprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   21
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblpdetail 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   3960
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   3
      Left            =   13560
      TabIndex        =   15
      Top             =   720
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   3
         Left            =   840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblpname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label lblprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   17
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblpdetail 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   3960
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   2
      Left            =   9240
      TabIndex        =   11
      Top             =   720
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   2
         Left            =   840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblpname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label lblprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   13
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblpdetail 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   3960
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   1
      Left            =   4920
      TabIndex        =   7
      Top             =   720
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   1
         Left            =   840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblpname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label lblprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   9
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblpdetail 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   3960
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   4215
      Begin VB.Label lblpdetail 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   3960
         Width           =   3615
      End
      Begin VB.Label lblprice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label lblpname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   0
         Left            =   840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Logoutbtn 
      BackColor       =   &H0080FFFF&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Home.frx":0000
      Left            =   12720
      List            =   "Home.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      Top             =   0
      Width           =   18735
   End
   Begin VB.Menu mnuprofile 
      Caption         =   "Profile"
      Begin VB.Menu mnuMyProfile 
         Caption         =   "My Profile"
      End
      Begin VB.Menu mnuMyOrder 
         Caption         =   "My Order"
      End
      Begin VB.Menu mnuMyCart 
         Caption         =   "My Cart"
      End
      Begin VB.Menu mnuinvoice 
         Caption         =   "Retrive Invoice"
      End
      Begin VB.Menu mnufeedback 
         Caption         =   "FeedBack"
      End
      Begin VB.Menu mnuSeller 
         Caption         =   "Switch To Seller"
      End
      Begin VB.Menu mnuAboutMe 
         Caption         =   "About Me"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim str As String
Dim i As Integer
Sub clear()
While i < 8
    Frame1(i).Visible = False
    i = i + 1
Wend
i = 0
End Sub

Sub display()
    Frame1(i).Visible = True
    lblpname(i).Caption = rs!ProductName
    lblprice(i).Caption = rs!PRICE & " RS"
    lblpdetail(i).Caption = rs!DETAIL
    Image1(i).Picture = LoadPicture(rs!PHOTO)
End Sub

Private Sub Cartbtn_Click()
CartForm.Show
End Sub

Private Sub Form_Load()
con.Open ("provider=microsoft.jet.oledb.4.0; data source=D:\Record.mdb;persist security info=false")
rs.Open ("select distinct category from Product"), con, adOpenDynamic, adLockPessimistic
rs1.Open ("select * from customer where CUSTOMERID=" + Customerlogin.txtcus_id.Text + " "), con, adOpenDynamic, adLockPessimistic
While Not rs.EOF
    Combo1.AddItem (rs!category)
    rs.MoveNext
Wend
clear
i = 0
End Sub

Private Sub combo1_click()
i = 0

If Combo1.Text <> " " Then
    rs.Close
    rs.Open ("select * from product where CATEGORY='" + Combo1.Text + "' "), con, adOpenDynamic, adLockPessimistic
    While Not rs.EOF
        display
        rs.MoveNext
        i = i + 1
    Wend
    clear
End If
End Sub

Private Sub Frame1_Click(Index As Integer)
OrderForm.Show
End Sub

Private Sub Image1_Click(Index As Integer)
OrderForm.Show
End Sub

Private Sub Logoutbtn_Click()
Customerlogin.txtcus_id.Text = ""
Customerlogin.txtpassword.Text = ""
Me.Hide
End Sub

Private Sub mnuAboutMe_Click()
AboutMe.Show
End Sub

Private Sub mnuFeedBack_Click()
Feedback.Show
End Sub

Private Sub mnuhelp_Click()
Help.Show
End Sub

Private Sub mnuinvoice_Click()
RetriveInvoice.Show
End Sub

Private Sub mnuMyCart_Click()
CartForm.Show
End Sub

Private Sub mnuMyOrder_Click()
MyOrder.Show

End Sub

Private Sub mnuMyProfile_Click()
rs1.Close
rs1.Open ("select * from customer where CUSTOMERID=" + Customerlogin.txtcus_id.Text + "  "), con, adOpenDynamic, adLockPessimistic
If rs.EOF Then
    CustomerProfile.txtid.Text = rs1!Customerid
    CustomerProfile.txtname.Text = rs1!Name
    CustomerProfile.txtaddress.Text = rs1!address
    CustomerProfile.txtphone.Text = rs1!contactno
    CustomerProfile.txtemail.Text = rs1!email
    CustomerProfile.Image1.Picture = LoadPicture(rs1!PHOTO)
End If
CustomerProfile.Show
End Sub

Private Sub mnuSeller_Click()
Customerlogin.txtcus_id.Text = ""
Customerlogin.txtpassword.Text = ""
Me.Hide
End Sub
