VERSION 5.00
Begin VB.Form Book 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Book"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   13230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
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
      Left            =   10800
      TabIndex        =   23
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
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
      Left            =   9120
      TabIndex        =   22
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
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
      Left            =   7560
      TabIndex        =   21
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
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
      Left            =   5880
      TabIndex        =   20
      Top             =   6240
      Width           =   1215
   End
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
      Left            =   9840
      TabIndex        =   19
      Top             =   4440
      Width           =   2415
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
      Left            =   9840
      TabIndex        =   18
      Top             =   3720
      Width           =   2415
   End
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
      Left            =   9840
      TabIndex        =   17
      Top             =   2880
      Width           =   2415
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
      Left            =   9840
      TabIndex        =   16
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox author 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   15
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox book_name 
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
      Left            =   3720
      TabIndex        =   14
      Top             =   3960
      Width           =   2655
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
      Left            =   3720
      TabIndex        =   13
      Top             =   3120
      Width           =   2655
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
      Left            =   3720
      TabIndex        =   12
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   6720
         TabIndex        =   3
         Top             =   1320
         Width           =   6375
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cost"
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
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   3120
            Width           =   1695
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Copies"
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
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Edition"
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
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Publisher"
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
            Height          =   615
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   6375
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
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   3600
            Width           =   1695
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Name of the book"
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
            Left            =   120
            TabIndex        =   6
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Subject"
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
            Left            =   120
            TabIndex        =   5
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ISBN Number"
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
            Left            =   120
            TabIndex        =   4
            Top             =   1200
            Width           =   1695
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Book Information"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Book"
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
    isbn_no.Text = ""
    subject.Text = ""
    book_name = ""
    author = ""
    publisher = ""
    edition = ""
    copies = ""
    cost = ""
End Sub

Private Sub Command2_Click()
If rs.State = 1 Then rs.Close

 If isbn_no.Text = "" Then
 MsgBox ("fill the ISBN "), vbInformation
 ElseIf book_name.Text = "" Then
 
 
 MsgBox ("fill  the book name "), vbInformation
  ElseIf subject.Text = "" Then
 
 MsgBox ("fill  the subject"), vbInformation


   ElseIf author.Text = "" Then

 MsgBox ("fill  the author name "), vbInformation
   ElseIf publisher.Text = "" Then
 
 MsgBox ("fill  the publisher Name "), vbInformation

    ElseIf edition.Text = "" Then
 
 MsgBox ("fill  the edition "), vbInformation

    ElseIf copies.Text = "" Then
 
 MsgBox ("enter the copies  "), vbInformation
  ElseIf cost.Text = "" Then
 
 MsgBox ("enter the cost  "), vbInformation

 Else

    SQL = "insert into book1(isbn_number,subject,book_name,author,publisher,edition,copies,cost)values('" + isbn_no.Text + "','" + subject.Text + "','" + book_name.Text + "','" + author.Text + "','" + publisher.Text + "','" + edition.Text + "','" + copies.Text + "','" + cost.Text + "') "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    MsgBox ("succesfully completed"), vbInformation0
    End If
End Sub

Private Sub Command3_Click()
    isbn_no.Text = ""
    subject.Text = ""
    book_name = ""
    author = ""
    publisher = ""
    edition = ""
    copies = ""
    cost = ""
End Sub

Private Sub Command4_Click()
  Unload Me
End Sub
Public Sub numerictest(obj As Object)
    If IsNumeric(obj.Text) Then
    Else
        MsgBox "Characters Not Allowed", vbCritical
        obj.Text = ""
    End If
End Sub

Private Sub copies_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(copies)
End Sub

Private Sub cost_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(cost)
End Sub

Private Sub edition_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(edition)
End Sub

Private Sub isbn_no_KeyUp(KeyCode As Integer, Shift As Integer)
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
