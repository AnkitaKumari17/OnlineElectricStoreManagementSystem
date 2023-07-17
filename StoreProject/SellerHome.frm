VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form SellerHome 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9495
   ClientLeft      =   -60
   ClientTop       =   585
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   18960
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2640
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\RECORD.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\RECORD.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from PRODUCT"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SellerHome.frx":0000
      Height          =   5895
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   10398
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "PRODUCT REPORT"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Viewallbtn 
      Caption         =   "View all record"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   5
      Top             =   8400
      Width           =   2175
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
      Left            =   12960
      TabIndex        =   4
      Text            =   "Select Category.................................."
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Refreshbtn 
      Caption         =   "# Refresh"
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
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Removeproductbtn 
      Caption         =   "Remove Product"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   2535
   End
   Begin VB.CommandButton Addproductbtn 
      Caption         =   "Add New Product"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Product"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   7200
      TabIndex        =   3
      Top             =   360
      Width           =   3630
   End
   Begin VB.Menu mnuSellerHome 
      Caption         =   "SELLER HOME"
      Begin VB.Menu mnuProfile 
         Caption         =   "My Profile"
      End
      Begin VB.Menu mnuCustomerMode 
         Caption         =   "Customer Mode"
      End
      Begin VB.Menu mnuFeedBack 
         Caption         =   "FeedBack"
      End
      Begin VB.Menu mnuAboutMe 
         Caption         =   "About Me"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "LOGOUT"
      Index           =   7
   End
End
Attribute VB_Name = "SellerHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim str As String

Private Sub Addproductbtn_Click()
AddNewProduct.Show
End Sub


Private Sub Form_Load()
con.Open ("provider=microsoft.jet.oledb.4.0; data source=D:\Record.mdb;persist security info=false")
rs.Open ("select distinct category from  Product"), con, adOpenDynamic, adLockPessimistic

Adodc1.RecordSource = "select CATEGORY,BRANDNAME,PRODUCTNAME,PRICE,STOCK,DETAIL,PHOTO from PRODUCT"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
While Not rs.EOF
    Combo1.AddItem (rs!category)
    rs.MoveNext
Wend

End Sub


Private Sub mnuAboutMe_Click()
AboutMe.Show
End Sub

Private Sub mnuCustomerMode_Click()
Me.Hide
Startup.Show
End Sub

Private Sub mnuFeedBack_Click()
Feedback.Show
End Sub

Private Sub mnuhelp_Click()
Help.Show
End Sub

Private Sub mnuLogout_Click(Index As Integer)
str = MsgBox("Are You Sure to LogOut", vbQuestion + vbYesNo)
If str = vbYes Then
    Me.Hide
    Startup.Show
End If
End Sub

Private Sub mnuProfile_Click()
'rs1.Close
rs1.Open ("select * from Seller where ID=" + SellerLogin.txtid.Text + "  "), con, adOpenDynamic, adLockPessimistic
If rs.EOF Then

MyProfile.txtsellerid.Text = rs1!id
MyProfile.txtname.Text = rs1!Name
MyProfile.txtaddress.Text = rs1!address
MyProfile.txtphone.Text = rs1!contactno
MyProfile.txtemail.Text = rs1!email
MyProfile.Image1.Picture = LoadPicture(rs1!PHOTO)
End If

MyProfile.Show

End Sub

Private Sub Refreshbtn_Click()
Adodc1.Refresh
While Not rs.EOF
    Combo1.AddItem (rs!category)
    rs.MoveNext
Wend
MsgBox "Table Updated"
End Sub

Private Sub Removeproductbtn_Click()
rs.Close
rs.Open ("select * from product"), con, adOpenDynamic, adLockPessimistic
If rs.BOF Then
    str = MsgBox("Sorry Record not found")
Else
    str = MsgBox("Do you want to remove the Product", vbQuestion + vbYesNo)
    If str = vbYes Then
        Adodc1.Recordset.Delete
        MsgBox ("Current Record deleted SuccessFully")
        Adodc1.Refresh
    End If
End If
End Sub


Private Sub combo1_click()
Adodc1.RecordSource = ("select * from product where CATEGORY='" + Combo1.Text + "' ")
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Viewallbtn_Click()
Adodc1.RecordSource = ("select * from product ")
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub
