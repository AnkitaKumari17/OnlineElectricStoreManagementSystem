VERSION 5.00
Begin VB.Form AboutMe 
   ClientHeight    =   8970
   ClientLeft      =   2280
   ClientTop       =   1170
   ClientWidth     =   17745
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   17745
   Begin VB.Image Image1 
      Height          =   2100
      Left            =   2520
      Picture         =   "AboutMe.frx":0000
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2220
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Magadh Mahila College (MMC)"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Under College-"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   5040
      Width           =   1920
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ankita,            Ankita kumari,       Rishika Raj"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6360
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developers:-"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   3360
      Width           =   1845
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"AboutMe.frx":451A
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   6360
      TabIndex        =   3
      Top             =   1680
      Width           =   8220
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About application"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   1080
      Width           =   2625
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "  About Developer"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2160
      TabIndex        =   1
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   7575
      Left            =   1800
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developer"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   1800
      Top             =   0
      Width           =   13455
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7935
      Left            =   1800
      Top             =   240
      Width           =   13455
   End
End
Attribute VB_Name = "AboutMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



