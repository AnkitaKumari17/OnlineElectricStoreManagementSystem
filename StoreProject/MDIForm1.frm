VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   18960
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Index           =   2
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Splash.Show
End Sub


Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuhelp_Click()
Help.Show
End Sub

