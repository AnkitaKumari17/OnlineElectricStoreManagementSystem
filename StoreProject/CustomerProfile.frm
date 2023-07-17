VERSION 5.00
Begin VB.Form CustomerProfile 
   ClientHeight    =   8325
   ClientLeft      =   2280
   ClientTop       =   825
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   15045
   Begin VB.TextBox txtid 
      BackColor       =   &H00DDFEFF&
      Height          =   450
      Left            =   8520
      TabIndex        =   12
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtaddress 
      BackColor       =   &H00DDFEFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8520
      TabIndex        =   10
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox txtemail 
      BackColor       =   &H00DDFEFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8520
      TabIndex        =   9
      Top             =   4320
      Width           =   3135
   End
   Begin VB.TextBox txtphone 
      BackColor       =   &H00DDFEFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8520
      TabIndex        =   8
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00DDFEFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8520
      TabIndex        =   7
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6480
      TabIndex        =   11
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6480
      TabIndex        =   6
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6480
      TabIndex        =   5
      Top             =   4320
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6480
      TabIndex        =   4
      Top             =   3600
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6480
      TabIndex        =   3
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Profile"
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
      Left            =   12720
      TabIndex        =   1
      Top             =   0
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2940
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   6120
      X2              =   14175
      Y1              =   1920
      Y2              =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Profile Detail"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Left            =   7920
      TabIndex        =   2
      Top             =   1200
      Width           =   4305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Profile"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   5040
      Width           =   2550
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7695
      Left            =   1440
      Top             =   600
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1440
      Top             =   0
      Width           =   13455
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   8295
      Left            =   1560
      Top             =   600
      Width           =   13455
   End
End
Attribute VB_Name = "CustomerProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
