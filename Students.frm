VERSION 5.00
Begin VB.Form Students 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Book Issue"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12720
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.CommandButton Command5 
         Caption         =   "Fine"
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
         Left            =   7440
         TabIndex        =   30
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Quit"
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
         Left            =   11280
         TabIndex        =   29
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Return"
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
         Left            =   9240
         TabIndex        =   28
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   6480
         TabIndex        =   14
         Top             =   720
         Width           =   6135
         Begin VB.TextBox Text4 
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
            Left            =   3000
            TabIndex        =   20
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox Text5 
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
            Height          =   465
            Left            =   3000
            TabIndex        =   19
            Top             =   870
            Width           =   2775
         End
         Begin VB.TextBox Text6 
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
            Height          =   420
            Left            =   3000
            TabIndex        =   18
            Top             =   1635
            Width           =   2775
         End
         Begin VB.TextBox Text7 
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
            Left            =   3000
            TabIndex        =   17
            Top             =   2160
            Width           =   2775
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3000
            TabIndex        =   16
            Top             =   2835
            Width           =   2775
         End
         Begin VB.TextBox Text9 
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
            Height          =   420
            Left            =   3000
            TabIndex        =   15
            Top             =   3435
            Width           =   2775
         End
         Begin VB.Label Label7 
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
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   2415
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
            Height          =   495
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fine"
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
            TabIndex        =   24
            Top             =   1560
            Width           =   2295
         End
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
            Height          =   495
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Date"
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
            TabIndex        =   22
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Issue Date"
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
            TabIndex        =   21
            Top             =   3480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   4275
         Left            =   240
         TabIndex        =   2
         Top             =   660
         Width           =   6135
         Begin VB.CommandButton Command1 
            Caption         =   "Click"
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
            Left            =   5040
            TabIndex        =   27
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text1 
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
            Left            =   2400
            TabIndex        =   8
            Top             =   1320
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   2535
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2040
            Width           =   2535
         End
         Begin VB.TextBox Text2 
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
            Height          =   375
            Left            =   2400
            TabIndex        =   5
            Top             =   2760
            Width           =   2535
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Click"
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
            Left            =   5040
            TabIndex        =   4
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox Text3 
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
            Left            =   2400
            TabIndex        =   3
            Top             =   3360
            Width           =   2535
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Memeber Name"
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
            TabIndex        =   13
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Address"
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
            TabIndex        =   12
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Book ID"
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
            TabIndex        =   11
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label Label5 
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
            TabIndex        =   10
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Book Name"
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
            TabIndex        =   9
            Top             =   3480
            Width           =   2415
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Return Books"
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
         Height          =   495
         Left            =   5520
         TabIndex        =   1
         Top             =   300
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Students"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cpy As Integer

Private Sub Command1_Click()
    If rs.State = 1 Then rs.Close
    SQL = "select * from student_book where member_name='" + Combo1.Text + "' "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    Text1.Text = (rs.Fields(2))
 
    
    If rs.State = 1 Then rs.Close
  
    SQL = "select * from student_book where member_name='" + Combo1.Text + "' "
      
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    While Not rs.EOF
        Combo2.AddItem (rs.Fields(3))
        rs.MoveNext
        
    Wend
   
    
    
End Sub

Private Sub Command2_Click()
     If rs.State = 1 Then rs.Close
    SQL = "select * from student_book where isbn_number='" + Combo2.Text + "' "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    Text2.Text = (rs.Fields(4))
    Text3.Text = (rs.Fields(5))
    Text4.Text = (rs.Fields(6))
    Text5.Text = (rs.Fields(7))
    Text7.Text = (rs.Fields(9))
    
    Text8.Text = (rs.Fields(10))
    Text9.Text = (rs.Fields(11))
    rs.Close
    
    
    
If rs.State = 1 Then rs.Close
SQL = "select * from  book1 where isbn_number='" + Combo2.Text + "' "
rs.Open SQL, con, adOpenDynamic, adLockOptimistic




cpy = rs.Fields(7)
   
   If (Text6.Text = DateDiff("d", Now, Text9.Text) < 0) Then
   
    Text6.Text = DateDiff("d", Now, Text9.Text)
    Else
    Text6.Text = "no fine"
    End If
    
    
End Sub

Private Sub Command3_Click()
    
    Dim new_cpy As Integer
    new_cpy = cpy + 1
    If rs.State = 1 Then rs.Close
    

SQL = "update book1 set copies='" & new_cpy & "' where  isbn_number='" & Combo2.Text & "' "
rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    If rs.State = 1 Then rs.Close
        
        
    SQL = "delete from student_book where member_name='" + Combo1.Text + "'"
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    MsgBox ("successfully returned")
    
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
    If rs.State = 1 Then rs.Close
    SQL = "insert into fine(mem_name,mem_address,book_id,subject,edition,issue_date)values('" + Combo1.Text + "','" + Text1.Text + "','" + Combo2.Text + "','" + Text2.Text + "','" + Text5.Text + "','" + Text9.Text + "') "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    MsgBox ("succesfully completed"), vbInformation0
End Sub

Private Sub Form_Load()
 If rs.State = 1 Then rs.Close
    SQL = "select * from student_book"
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
   ' rs.MoveFirst
    While Not rs.EOF
        Combo1.AddItem (rs.Fields(1))
        rs.MoveNext
        
    Wend
    

    
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
    Case vbKey0 To vbKey9
    KeyAscii = 0
    Beep
    MsgBox "Numeric Not Allowed", vbCritical
    End Select
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
    Case vbKey0 To vbKey9
    KeyAscii = 0
    Beep
    MsgBox "Numeric Not Allowed", vbCritical
    End Select
End Sub
Public Sub numerictest(obj As Object)
    If IsNumeric(obj.Text) Then
    Else
        MsgBox "Characters Not Allowed", vbCritical
        obj.Text = ""
    End If
End Sub
Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(Text5)
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(Text6)
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(Text7)
End Sub
