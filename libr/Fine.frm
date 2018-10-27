VERSION 5.00
Begin VB.Form Fine 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fine Information"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   12780
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
      Left            =   9120
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox issue_date 
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
      Left            =   9240
      TabIndex        =   8
      Top             =   3000
      Width           =   2775
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
      Left            =   3000
      TabIndex        =   7
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   6600
         TabIndex        =   3
         Top             =   1560
         Width           =   6015
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
            Height          =   495
            Left            =   2640
            TabIndex        =   17
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox fine_amt 
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
            Left            =   2640
            TabIndex        =   13
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label6 
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
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fine Amount"
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
            TabIndex        =   12
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Return Date"
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
            TabIndex        =   6
            Top             =   1560
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   6255
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
            Height          =   450
            Left            =   2760
            TabIndex        =   19
            Top             =   1800
            Width           =   2775
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
            Left            =   2760
            TabIndex        =   15
            Top             =   480
            Width           =   2775
         End
         Begin VB.CommandButton Command5 
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
            Height          =   375
            Left            =   4800
            TabIndex        =   14
            Top             =   1080
            Width           =   1095
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
            Height          =   420
            Left            =   2760
            TabIndex        =   10
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label8 
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
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label2 
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
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label4 
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
            Left            =   240
            TabIndex        =   5
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Member Name"
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
            TabIndex        =   4
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fine Details"
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
         Height          =   735
         Left            =   5400
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Fine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
mem_id.Text = " "
mem_name.Text = " "
mem_local.Text = " "
mem_per.Text = " "
mem_phone.Text = " "
mem_type.Text = " "
mem_gender.Text = " "
End Sub

Private Sub Command2_Click()
 If rs.State = 1 Then rs.Close
    SQL = "insert into fine(fine_id,mem_name,mem_local,mem_per,mem_phone,mem_type,mem_gender)values('" + mem_id.Text + "','" + mem_name.Text + "','" + mem_local.Text + "','" + mem_per.Text + "','" + mem_phone.Text + "','" + mem_type.Text + "','" + mem_gender.Text + "') "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    MsgBox ("succesfully completed"), vbInformation

End Sub

Private Sub Command3_Click()
mem_id.Text = " "
mem_name.Text = " "
mem_local.Text = " "
mem_per.Text = " "
mem_phone.Text = " "
mem_type.Text = " "
mem_gender.Text = " "
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Dim aa As Date

If rs.State = 1 Then rs.Close
    SQL = "select * from student_book where member_name='" + Combo1.Text + "' "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    Text3.Text = (rs.Fields(2))
    Text1.Text = (rs.Fields(3))
    book_name.Text = (rs.Fields(4))
    Text2.Text = (rs.Fields(7))
    issue_date.Text = (rs.Fields(11))
    rs.Close
    
aa = DateDiff("d", Now, issue_date.Text)
    
    If (aa < 0) Then
    fine_amt.Text = Abs(DateDiff("d", Now, issue_date.Text))
    Else
    fine_amt.Text = 0
    
    End If
    
End Sub

Private Sub Command6_Click()

End Sub


Private Sub fine_amt_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(fine_amt)
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
    rs.Close
    

'fine_amt.Text = currentDate.AddDays(1)
End Sub

Public Sub numerictest(obj As Object)
    If IsNumeric(obj.Text) Then
    Else
        MsgBox "Characters Not Allowed", vbCritical
        obj.Text = ""
    End If
End Sub

Private Sub issue_date_Change()
   ' Call numerictest(issue_date)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(Text1)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(Text2)
End Sub
