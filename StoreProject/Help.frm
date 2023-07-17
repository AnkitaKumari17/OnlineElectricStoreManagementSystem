VERSION 5.00
Begin VB.Form Help 
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   2220
   ClientTop       =   465
   ClientWidth     =   16740
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   16740
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "              If there are come any error then close the program and restart the application."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   6960
      TabIndex        =   5
      Top             =   5520
      Width           =   7455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"Help.frx":0000
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2490
      Left            =   6960
      TabIndex        =   4
      Top             =   2880
      Width           =   7320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "How to use"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6960
      TabIndex        =   3
      Top             =   2160
      Width           =   2145
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't Worry! You come exect page. Here we  give you some point to handle this application. So Let's Start............! "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      TabIndex        =   2
      Top             =   1200
      Width           =   7335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   4200
      Width           =   1860
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   7455
      Left            =   2760
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HELP"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12960
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2760
      Top             =   0
      Width           =   12015
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8175
      Left            =   3000
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
