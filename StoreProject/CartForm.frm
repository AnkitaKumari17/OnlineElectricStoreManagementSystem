VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form CartForm 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   300
   ClientTop       =   645
   ClientWidth     =   18885
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   18885
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "CartForm.frx":0000
      Height          =   5535
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   9763
      _Version        =   393216
      AllowUpdate     =   0   'False
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   24
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Item In Cart"
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
   Begin VB.CommandButton Checkoutbtn 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Check Out"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton Remitembtn 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remove item"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3360
      Top             =   8400
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   "select * from cart"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4080
      Top             =   6600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "CART"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lbltotalamt 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   15840
      TabIndex        =   7
      Top             =   8520
      Width           =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL AMOUNT"
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
      Left            =   13920
      TabIndex        =   6
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Cart"
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
      Left            =   8520
      TabIndex        =   1
      Top             =   840
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Cart"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   17040
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "CartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim result As Integer

Sub display()

rs.Close
rs.Open ("select * from CART where  CUSTOMERID='" + Customerlogin.txtcus_id.Text + "' "), con, adOpenDynamic, adLockPessimistic
Adodc2.RecordSource = ("select * from cart where CUSTOMERID='" + Customerlogin.txtcus_id.Text + "' ")
Adodc2.Refresh
Adodc2.Caption = Adodc2.RecordSource
Set DataGrid1.DataSource = Adodc2

End Sub

Private Sub Checkoutbtn_Click()
PaymentForm.Show

End Sub

Private Sub Command1_Click()
Me.Hide
Home.Show
End Sub

Private Sub Form_Load()
If con.State = 1 Then
    con.Close
End If

con.Open ("provider=microsoft.jet.oledb.4.0; data source=D:\Record.mdb;persist security info=false")
rs.Open ("select * from CART where CUSTOMERID='" + Customerlogin.txtcus_id.Text + "' "), con, adOpenDynamic, adLockPessimistic

display
result = 0
While Not rs.EOF
    result = rs!amount + result
    rs.MoveNext
Wend
    
lbltotalamt.Caption = result

If result = 0 Then
    Checkoutbtn.Enabled = False
End If

End Sub

Private Sub Remitembtn_Click()
If rs.BOF Then
    str = MsgBox("Sorry Record not found")
Else
    str = MsgBox("Do you want to remove the Product", vbQuestion + vbYesNo)
    If str = vbYes Then
'        rs.Delete
 '       rs.Update
        Adodc2.Recordset.Delete
        MsgBox ("Current Record deleted SuccessFully")
        Adodc2.Refresh
'        rs.MoveNext
    End If
End If
Adodc2.Refresh
MsgBox "Table Updated"
While Not rs.EOF
    result = rs!amount + result
    rs.MoveNext
Wend
    
lbltotalamt.Caption = result

End Sub
