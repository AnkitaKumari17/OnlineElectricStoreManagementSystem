VERSION 5.00
Begin VB.Form PaymentForm 
   ClientHeight    =   8505
   ClientLeft      =   645
   ClientTop       =   -60
   ClientWidth     =   17865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   17865
   Begin VB.CommandButton Savebtn 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtdaddress 
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
      Left            =   8880
      TabIndex        =   11
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton Nextbtn 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Next>>"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Cancelorderbtn 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cancel Order"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   6120
      TabIndex        =   1
      Top             =   3240
      Width           =   9495
      Begin VB.OptionButton Option4 
         Caption         =   "Net Banking"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   3240
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Debit/ Credit Card"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   2520
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "EMI"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cash On Delivery"
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
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Payment Method"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Label lblamount 
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
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   9000
      TabIndex        =   13
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAYBLE AMOUNT"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   12
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DELIVERY ADDRESS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   10
      Top             =   960
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   3000
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   3000
      TabIndex        =   9
      Top             =   5400
      Width           =   2235
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Type"
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
      Left            =   11520
      TabIndex        =   0
      Top             =   120
      Width           =   1860
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   7815
      Left            =   2280
      Top             =   600
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   2280
      Top             =   0
      Width           =   13695
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7815
      Left            =   5040
      Top             =   600
      Width           =   10935
   End
End
Attribute VB_Name = "PaymentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim str As String
Dim item As Integer
Sub updateorderhistory()
rs.Close
rs.Open ("select * from cart where CUSTOMERID='" + Customerlogin.txtcus_id.Text + "' ")
'rs3.Close
rs3.Open ("select * from orderhistory "), con, adOpenDynamic, adLockPessimistic
rs3.AddNew
rs3.Fields(0).Value = rs!ProductName
rs3.Fields(1).Value = rs!BrandName
rs3.Fields(2).Value = rs!Productdetail
rs3.Fields(3).Value = rs!quantity
rs3.Fields(4).Value = rs!Rate
rs3.Fields(5).Value = rs!amount
rs3.Fields(6).Value = rs!Customerid
rs3.Fields(7).Value = rs!orderid
rs3.Update

End Sub

Sub remove()
rs3.Close
rs3.Open ("select * from orderhistory where PRODUCTNAME='" + OrderForm.lblpname.Caption + "' "), con, adOpenDynamic, adLockPessimistic
'rs2.Close
Dim sql As String
sql = "select * from product where PRICE='" + Replace(OrderForm.lblprice.Caption, " RS", "") + "' "
rs2.Open (sql), con, adOpenDynamic, adLockPessimistic
item = Val(rs2!stock) - rs3!quantity
rs2.Fields(4).Value = item
rs2.Update
rs2.Close

End Sub

Private Sub Cancelorderbtn_Click()
Me.Hide
CartForm.Show
End Sub

Private Sub Form_Load()
con.Open ("provider=microsoft.jet.OLEDB.4.0;data source=D:\Record.mdb;persist security info=false")
rs.Open ("select * from cart where CUSTOMERID='" + Customerlogin.txtcus_id.Text + "' "), con, adOpenDynamic, adLockPessimistic
rs1.Open ("select * from payment"), con, adOpenDynamic, adLockPessimistic
lblamount.Caption = CartForm.lbltotalamt & "RS"
End Sub

Private Sub Nextbtn_Click()
If txtdaddress.Text <> " " Then

If Option3.Value = True Then
   MsgBox "NOT Available for Now Try After Some Time "
   'Payform.Show
Else
    If Option2.Value = True Then
        MsgBox "No EMI available for this product"
        Exit Sub
    Else
        If Option4.Value = True Then
            MsgBox "No Net working available for this product"
            Exit Sub
        Else
            rs1.AddNew
            rs1.Fields(0).Value = Customerlogin.txtcus_id.Text
            rs1.Fields(2).Value = Val(lblamount.Caption)
            rs1.Fields(3).Value = txtdaddress.Text
            rs1.Fields(4).Value = Now()
            rs1.Fields(5).Value = Option1.Caption
            
            str = MsgBox(" Order successfully placed .Your Invoice id is '" & rs1!INVOICENO & "' ")
            rs1.Update
            
            updateorderhistory
            
            rs.Close
            rs.Open ("select * from cart where CUSTOMERID='" + Customerlogin.txtcus_id.Text + "' ")
            While Not rs.EOF
                rs.Delete
                rs.Update
                CartForm.Adodc2.Refresh
                rs.MoveNext
            Wend
            CartForm.lbltotalamt.Caption = 0
            Me.Hide
            
            remove
        End If
    End If
End If
Else
    MsgBox " Please fill the address "
End If

End Sub


