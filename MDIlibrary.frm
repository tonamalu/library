VERSION 5.00
Begin VB.MDIForm library 
   Appearance      =   0  'Flat
   BackColor       =   &H00000040&
   Caption         =   "Library Management"
   ClientHeight    =   9105
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   19830
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIlibrary.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuregister 
      Caption         =   "Register"
      Index           =   0
      Begin VB.Menu mnustudent 
         Caption         =   "Student Register"
         Index           =   1
         Shortcut        =   ^R
      End
      Begin VB.Menu mnumember 
         Caption         =   "Member Information"
         Index           =   2
         Begin VB.Menu mnusearch 
            Caption         =   "Search Member"
            Index           =   4
            Shortcut        =   ^F
         End
         Begin VB.Menu mnumemberdetails 
            Caption         =   "Members Details"
            Index           =   3
            Shortcut        =   ^J
         End
      End
   End
   Begin VB.Menu mnubook 
      Caption         =   "Book"
      Index           =   0
      Begin VB.Menu mnubookreg 
         Caption         =   "Book Register"
         Index           =   1
         Shortcut        =   ^B
      End
      Begin VB.Menu mnubookinfo 
         Caption         =   "Book Information"
         Index           =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuentire 
         Caption         =   "Entire Library Books"
         Index           =   4
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Fine"
      Index           =   0
      Begin VB.Menu mnufineinfo 
         Caption         =   "Fine Information"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnustudentfine 
         Caption         =   "Entire Students Fine"
         Index           =   3
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnulibrarian 
      Caption         =   "Librarian Information"
      Index           =   0
      Begin VB.Menu mnuchangepass 
         Caption         =   "Change password"
         Index           =   2
         Shortcut        =   ^P
      End
      Begin VB.Menu mnulibreg 
         Caption         =   "Librarian Registration"
         Index           =   1
         Shortcut        =   ^H
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnustudentbook 
      Caption         =   "Student Book"
      Index           =   0
      Begin VB.Menu mnulibrary 
         Caption         =   "Library Text"
         Index           =   1
         Begin VB.Menu mnustudents 
            Caption         =   "issue Students"
            Index           =   3
         End
      End
      Begin VB.Menu mnuissue 
         Caption         =   "Return Book"
         Index           =   2
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
      Index           =   0
   End
End
Attribute VB_Name = "library"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Call connect
library.Width = 15109
library.Height = 11010
End Sub



Private Sub mnubookinfo_Click(Index As Integer)
Book_details.Show
End Sub

Private Sub mnubookreg_Click(Index As Integer)
Book.Show
End Sub

Private Sub mnuchangepass_Click(Index As Integer)
changepass.Show
End Sub

Private Sub mnuentire_Click(Index As Integer)
LibraryBooks.Show
End Sub

Private Sub mnuexit_Click(Index As Integer)
End
End Sub

Private Sub mnufineinfo_Click(Index As Integer)
Fine.Show
End Sub

Private Sub mnuissue_Click(Index As Integer)
Students.Show
End Sub

Private Sub mnulibreg_Click(Index As Integer)
Librarian.Show
End Sub

Private Sub mnulogin1_Click(Index As Integer)
login.Show
library.Visible = False

End Sub

Private Sub mnumemberdetails_Click(Index As Integer)
member_full_details.Show
End Sub

Private Sub mnusearch_Click(Index As Integer)
Member_Details.Show
End Sub

Private Sub mnustudent_Click(Index As Integer)
Member.Show
End Sub

Private Sub mnustudentfine_Click(Index As Integer)
Fine_details.Show
End Sub

Private Sub mnustudents_Click(Index As Integer)
Form1.Show
End Sub
