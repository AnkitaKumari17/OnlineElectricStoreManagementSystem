VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form RetriveInvoice 
   Caption         =   "RETRIVE INVOICE FORM"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   Begin VB.CommandButton Retriveinvoicebtn 
      BackColor       =   &H00C0C0FF&
      Caption         =   "RETRIVE INVOICE"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1440
      Top             =   2640
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
      RecordSource    =   "select *  from orderhistory"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Exitbtn 
      BackColor       =   &H00C0FFFF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Cambria Math"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox txtinvoice 
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "TAX INVOICE"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   11775
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "RetriveInvoice.frx":0000
         Height          =   4335
         Left            =   240
         TabIndex        =   27
         Top             =   3240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7646
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ORDER LIST"
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
      Begin VB.CommandButton Printbtn 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Cambria Math"
            Size            =   11.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         Caption         =   "Buyer Detail"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   9615
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   32
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   31
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer ID:-"
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
            Left            =   5280
            TabIndex        =   30
            Top             =   480
            Width           =   1440
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                      "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6240
            TabIndex        =   23
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                    "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6720
            TabIndex        =   22
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                    "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1320
            TabIndex        =   21
            Top             =   1920
            Width           =   900
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "             "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6480
            TabIndex        =   20
            Top             =   1920
            Width           =   585
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "             "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1560
            TabIndex        =   19
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "            "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1560
            TabIndex        =   18
            Top             =   480
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail:-"
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
            Left            =   5280
            TabIndex        =   17
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact no:-"
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
            Left            =   5280
            TabIndex        =   16
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "District:-"
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
            Left            =   240
            TabIndex        =   15
            Top             =   1920
            Width           =   885
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pin code:-"
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
            Left            =   5280
            TabIndex        =   14
            Top             =   1920
            Width           =   1020
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "State:-"
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
            Left            =   240
            TabIndex        =   13
            Top             =   1440
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:-"
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
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:-"
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
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Label Label28 
         Caption         =   "Label28"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   29
         Top             =   7680
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "Label21"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   7800
         Width           =   1935
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authorised Signatory"
         BeginProperty Font 
            Name            =   "Myanmar Text"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7680
         TabIndex        =   26
         Top             =   8280
         Width           =   1830
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payable Amount:-"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   25
         Top             =   7680
         Width           =   2220
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode/Terms of Payment:-"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   7800
         Width           =   2490
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   10080
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:-"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9360
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8160
         TabIndex        =   8
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No:-"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6840
         TabIndex        =   7
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Invoice number :-"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   2595
   End
End
Attribute VB_Name = "RetriveInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim str As String

Private Sub Exitbtn_Click()
Frame1.Visible = False
txtinvoice.Text = ""
Me.Hide
End Sub

Private Sub Form_Load()
con.Open ("provider=microsoft.jet.oledb.4.0;data source=D:\Record.mdb;persist security info=false")
rs.Open ("select * from payment"), con, adOpenDynamic, adLockPessimistic
rs1.Open ("select * from customer"), con, adOpenDynamic, adLockPessimistic
Adodc1.RecordSource = ("SELECT ORDERID,BRANDNAME,PRODUCTNAME,QUANTITY,RATE,AMOUNT,CUSTOMERID FROM ORDERHISTORY where CUSTOMERID=" + Customerlogin.txtcus_id.Text + " ")
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
Set DataGrid1.DataSource = Adodc1
Frame1.Visible = False
End Sub

Private Sub Printbtn_Click()
Me.PrintForm
End Sub

Private Sub Retriveinvoicebtn_Click()
rs.Close
rs.Open ("select * from PAYMENT where INVOICENO=" & txtinvoice.Text & " "), con, adOpenDynamic, adLockPessimistic

If rs.EOF Then
    str = MsgBox("Invalid Details", vbExclamation + vbDefaultButton1)
Else
    Label6.Caption = rs!Date
    Label4.Caption = rs!INVOICENO
    Label28.Caption = rs!PAIDAMOUNT
    Label21.Caption = rs!PAYMENTMODE
    rs1.Close
    rs1.Open ("select * from customer where CUSTOMERID=" + Customerlogin.txtcus_id.Text + " "), con, adOpenDynamic, adLockPessimistic
    If Not rs.EOF Then
        Label14.Caption = rs1!Name
        Label20.Caption = rs1!Customerid
        Label22.Caption = rs1!address
        Label15.Caption = rs1!State
        Label17.Caption = rs1!DISTRICT
        Label18.Caption = rs1!contactno
        Label19.Caption = rs1!email
        Label16.Caption = rs1!PINCODE
    Else
        str = MsgBox("Buyer Address Not Retrived", vbExclamation + vbDefaultButton1)
    End If
        Adodc1.RecordSource = ("select * from  where CUSTOMERID='" + Customerlogin.txtcus_id.Text + "' ")
'        Adodc1.Refresh
        Frame1.Visible = True
    End If
End Sub
