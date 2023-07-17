VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form AddNewProduct 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9300
   ClientLeft      =   315
   ClientTop       =   660
   ClientWidth     =   18960
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "AddNewProduct.frx":0000
   ScaleHeight     =   9300
   ScaleWidth      =   18960
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13080
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Uploadphotobtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Upload Photo"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtpdetail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   15
      Top             =   5880
      Width           =   3135
   End
   Begin VB.TextBox txtstock 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   14
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox txtpprice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton Cancelbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8160
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "AddNewProduct.frx":0342
      Left            =   6600
      List            =   "AddNewProduct.frx":0355
      TabIndex        =   10
      Text            =   "SELECT"
      Top             =   2640
      Width           =   3135
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
      ItemData        =   "AddNewProduct.frx":0384
      Left            =   6600
      List            =   "AddNewProduct.frx":03A0
      TabIndex        =   9
      Text            =   "SELECT"
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtpname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   3360
      Width           =   3135
   End
   Begin VB.CommandButton Resetbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Addproductbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add Product "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      FillColor       =   &H00000080&
      Height          =   2415
      Left            =   14760
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT DETAIL:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL QUANTITY:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT PRICE:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   4320
      Width           =   1770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT NAME:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   3480
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CATAGORIES:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BRAND NAME:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Shape Shape2 
      Height          =   9255
      Left            =   240
      Top             =   0
      Width           =   18615
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   360
      X2              =   720
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   360
      X2              =   720
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Product"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15960
      TabIndex        =   0
      Top             =   120
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   240
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "AddNewProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim str1 As String

Sub clear()
Combo1.Text = ""
Combo2.Text = ""
txtpname.Text = ""
txtpprice.Text = ""
txtstock.Text = ""
txtpdetail.Text = ""
Image1.Picture = LoadPicture("")
End Sub

Private Sub Addproductbtn_Click()
If Combo1.Text = "" Or Combo2.Text = "" Or txtpname.Text = "" Or txtpprice.Text = "" Or txtstock = "" Or txtpdetail.Text = "" Or Image1.Picture = LoadPicture("") Then
    str = MsgBox("Some details are empty", vbExclamation + vbDefaultButton1)
Else
    rs.AddNew
    rs.Fields(0).Value = Combo1.Text
    rs.Fields(1).Value = Combo2.Text
    rs.Fields(2).Value = txtpname.Text
    rs.Fields(3).Value = txtpprice.Text
    rs.Fields(4).Value = txtstock.Text
    rs.Fields(5).Value = txtpdetail.Text
    rs.Fields(6).Value = str1
    rs.Update
    str = MsgBox("Product added successfully", vbInformation + vbDefaultButton1)
    clear
End If
       
End Sub

Private Sub Cancelbtn_Click()
clear
Me.Hide
End Sub

Private Sub Form_Load()
con.Open ("provider=microsoft.jet.oledb.4.0; data source=D:\Record.mdb;persist security info=false")
rs.Open ("select *from PRODUCT"), con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Resetbtn_Click()
clear
End Sub

Private Sub Uploadphotobtn_Click()
CommonDialog1.Filter = "jpg|*.jpg|jpeg|*.jpeg"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
str1 = CommonDialog1.FileName
Image1.Picture = LoadPicture(str1)
End Sub
