VERSION 5.00
Begin VB.Form Feedback 
   ClientHeight    =   9390
   ClientLeft      =   465
   ClientTop       =   645
   ClientWidth     =   18690
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   18690
   Begin VB.CommandButton Submitbtn 
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
      Height          =   615
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Cancelbtn 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7800
      TabIndex        =   9
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   8
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   7
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Explain"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   4320
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   3480
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter email"
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
      Left            =   5760
      TabIndex        =   3
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Co. Email"
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
      Left            =   5760
      TabIndex        =   2
      Top             =   1680
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FeedBack"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12480
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FeedBack"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   4680
      Width           =   2565
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   8055
      Left            =   1320
      Top             =   600
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8055
      Left            =   1920
      Top             =   600
      Width           =   14055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1320
      Top             =   0
      Width           =   14655
   End
End
Attribute VB_Name = "Feedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancelbtn_Click()
Me.Hide
End Sub
