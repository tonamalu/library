VERSION 5.00
Begin VB.Form Book_details 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox cost 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   22
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox copies 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   21
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox isbn_no 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Book details"
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9840
         TabIndex        =   27
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7800
         TabIndex        =   26
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         TabIndex        =   25
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   24
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   23
         Top             =   5400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   6600
         TabIndex        =   6
         Top             =   1320
         Width           =   6015
         Begin VB.TextBox edition 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2760
            TabIndex        =   20
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox publisher 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2760
            TabIndex        =   19
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cost"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Copies"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   13
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Edition"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Publisher"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   6255
         Begin VB.TextBox author 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            TabIndex        =   18
            Top             =   2640
            Width           =   2295
         End
         Begin VB.TextBox b_name 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            TabIndex        =   17
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox subject 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            TabIndex        =   16
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Author"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   10
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Name of the Book"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   9
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Subject"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ISBN Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   12375
         Begin VB.CommandButton Command1 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8880
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4920
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Book Number"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "Book_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub author_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    KeyAscii = 0
    Beep
    MsgBox "Numeric Not Allowed", vbCritical
    End Select
End Sub

Private Sub Command1_Click()
If rs.State = 1 Then rs.Close
    SQL = "select * from book1 where book_number='" + Combo1.Text + "' "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    isbn_no.Text = (rs.Fields(1))
    subject.Text = (rs.Fields(2))
    b_name.Text = (rs.Fields(3))
    author.Text = (rs.Fields(4))
    publisher.Text = (rs.Fields(5))
    edition.Text = (rs.Fields(6))
    copies.Text = (rs.Fields(7))
    cost.Text = (rs.Fields(8))
    rs.Close
End Sub

Private Sub Command2_Click()
    LibraryBooks.Show
End Sub

Private Sub Command3_Click()
 If rs.State = 1 Then rs.Close
    SQL = "update book1 set isbn_number='" + isbn_no.Text + "',subject='" + subject.Text + "',book_name='" + b_name.Text + "',author='" + author.Text + "',publisher='" + publisher.Text + "',edition='" + edition.Text + "',copies='" + copies.Text + "',cost='" + cost.Text + "' where book_number='" + Combo1.Text + "' "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    MsgBox "Data Updated Successfully", vbInformation

End Sub

Private Sub Command4_Click()
If MsgBox("Do You Want to Delete the Customer ..?", vbQuestion + vbYesNo) = vbYes Then
    Command5.Enabled = True
    End If
End Sub

Private Sub Command5_Click()
If rs.State = 1 Then rs.Close
    Command5.Enabled = False
    If MsgBox("Do you permanantly delete?", vbQuestion + vbYesNo) = vbYes Then
    SQL = "delete from book1 where book_number='" + Combo1.Text + "'"
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    Else
    MsgBox "Deleted...!", vbInformation
    End If
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub copies_Change()
    Call numerictest(copies)
End Sub

Private Sub cost_Change()
    Call numerictest(cost)
End Sub

Private Sub edition_Change()
    Call numerictest(edition)
End Sub

Private Sub Form_Load()
 If rs.State = 1 Then rs.Close
    SQL = "select * from book1"
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    While Not rs.EOF
        Combo1.AddItem (rs.Fields(0))
        rs.MoveNext
    Wend
    rs.Close
    Command5.Enabled = False
End Sub
Public Sub numerictest(obj As Object)
    'If IsNumeric(obj.Text) Then
    'Else
    '    MsgBox "Characters Not Allowed", vbCritical
    '    obj.Text = ""
    'End If
End Sub

Private Sub isbn_no_Change()
    Call numerictest(isbn_no)
End Sub


Private Sub subject_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9
    KeyAscii = 0
    Beep
    MsgBox "Numeric Not Allowed", vbCritical
    End Select
End Sub
