VERSION 5.00
Begin VB.Form OrderForm 
   BackColor       =   &H8000000E&
   ClientHeight    =   9270
   ClientLeft      =   1005
   ClientTop       =   645
   ClientWidth     =   15750
   ForeColor       =   &H00808000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15750
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   14280
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton Addtocartbtn 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADD TO CART"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton Buynowbtn 
      BackColor       =   &H00C0FFFF&
      Caption         =   "BUY NOW"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label lblprice 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   8040
      TabIndex        =   13
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   8040
      TabIndex        =   12
      Top             =   4920
      Width           =   885
   End
   Begin VB.Label lblpdetail 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   11
      Top             =   3960
      Width           =   5055
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT DETAIL"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   8040
      TabIndex        =   10
      Top             =   3480
      Width           =   2700
   End
   Begin VB.Label lblbname 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   9
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label lblpname 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   8
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Label lbltotalprice 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7800
      TabIndex        =   6
      Top             =   7080
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Quantity"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4440
      TabIndex        =   5
      Top             =   7080
      Width           =   1875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BRAND NAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   8040
      TabIndex        =   3
      Top             =   2160
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT NAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   8040
      TabIndex        =   0
      Top             =   840
      Width           =   2505
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   960
      Width           =   6495
   End
End
Attribute VB_Name = "OrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim str As String
Dim i As Integer
Sub display()
    Image1.Picture = LoadPicture(rs!PHOTO)
    lblpname.Caption = rs!ProductName
    lblprice.Caption = rs!PRICE & " RS"
    lblpdetail.Caption = rs!DETAIL
    lblbname.Caption = rs!BrandName
End Sub

Private Sub Addtocartbtn_Click()
rs1.Close
rs1.Open ("select * from CART"), con, adOpenDynamic, adLockPessimistic
rs2.Close
rs2.Open ("select * from customer"), con, adOpenDynamic, adLockPessimistic
rs.Close
rs.Open ("select * from product where PRICE='" + Replace(OrderForm.lblprice.Caption, " RS", "") + "' "), con, adOpenDynamic, adLockPessimistic
If Val(rs!stock) = 0 Then
    MsgBox "stock not avilable"
    Text1.Text = ""
    Me.Hide
ElseIf Val(rs!stock) < Text1.Text Then
    MsgBox ("only " & rs!stock & " quantity avilable")
Else
    rs1.AddNew
    rs1.Fields(0).Value = lblpname.Caption
    rs1.Fields(1).Value = lblbname.Caption
    rs1.Fields(2).Value = lblpdetail.Caption
    rs1.Fields(3).Value = Text1.Text
    rs1.Fields(4).Value = Val(lblprice.Caption)
    rs1.Fields(5).Value = Val(lbltotalprice.Caption)
    rs1.Fields(6).Value = Customerlogin.txtcus_id.Text
    rs1.Update
    str = MsgBox("Product added to cart successfully", vbExclamation + vbInformation)
    CartForm.Adodc2.Refresh
End If

End Sub

Private Sub Command1_Click()
rs.MovePrevious
Text1.Text = ""
If rs.BOF Then
    rs.MoveLast
    display
Else
    display
End If

End Sub

Private Sub Command2_Click()
rs.MoveNext
Text1.Text = ""
If rs.EOF Then
    rs.MoveFirst
    display
Else
    display
End If

End Sub

Private Sub Form_Load()
If con.State = 1 Then
con.Close
End If

con.Open ("provider=microsoft.jet.oledb.4.0; data source=D:\Record.mdb;persist security info=false")
rs.Open ("select * from  product"), con, adOpenDynamic, adLockPessimistic

rs.Close
rs.Open ("select * from  product where category ='" + Home.Combo1.Text + "' "), con, adOpenDynamic, adLockPessimistic
display

rs1.Open ("select * from CART"), con, adOpenDynamic, adLockPessimistic
rs2.Open ("select * from customer"), con, adOpenDynamic, adLockPessimistic

lbltotalprice.Caption = Val(Text1.Text) * Val(lblprice.Caption)
End Sub

Private Sub Text1_Change()
lbltotalprice.Caption = Val(Text1.Text) * Val(lblprice.Caption) & "RS"
End Sub
