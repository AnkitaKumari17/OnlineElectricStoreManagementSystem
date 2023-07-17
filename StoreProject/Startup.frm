VERSION 5.00
Begin VB.Form Startup 
   AutoRedraw      =   -1  'True
   Caption         =   "STARTUP FORM"
   ClientHeight    =   8355
   ClientLeft      =   450
   ClientTop       =   810
   ClientWidth     =   18075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Startup.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   18075
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Sellerbtn 
      Caption         =   "SELLER LOGIN"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   5160
      Width           =   2775
   End
   Begin VB.CommandButton Quitbtn 
      BackColor       =   &H80000015&
      Caption         =   "QUIT"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   6120
      Width           =   2775
   End
   Begin VB.CommandButton Signupbtn 
      BackColor       =   &H80000015&
      Caption         =   "SIGNUP"
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
      Left            =   3840
      TabIndex        =   1
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Loginbtn 
      BackColor       =   &H80000015&
      Caption         =   "LOGIN"
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
      Left            =   3840
      TabIndex        =   0
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Electrical Store Management System"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   10170
   End
End
Attribute VB_Name = "Startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Loginbtn_Click()
Customerlogin.Show
End Sub


Private Sub Quitbtn_Click()
End
End Sub

Private Sub Sellerbtn_Click()
CreateSeller.Show

End Sub

Private Sub Signupbtn_Click()
CreateCustomer.Show
End Sub

