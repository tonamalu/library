VERSION 5.00
Begin VB.Form Member_Details 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Member Details"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox mem_gender 
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
      Left            =   5040
      TabIndex        =   18
      Top             =   6000
      Width           =   3735
   End
   Begin VB.TextBox mem_type 
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
      Left            =   5040
      TabIndex        =   17
      Top             =   5280
      Width           =   3735
   End
   Begin VB.TextBox mem_phone 
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
      Left            =   5040
      TabIndex        =   16
      Top             =   4560
      Width           =   3735
   End
   Begin VB.TextBox mem_per 
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
      Left            =   5040
      TabIndex        =   15
      Top             =   3840
      Width           =   3735
   End
   Begin VB.TextBox mem_local 
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
      Left            =   5040
      TabIndex        =   14
      Top             =   3240
      Width           =   3735
   End
   Begin VB.TextBox mem_name 
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
      Left            =   5040
      TabIndex        =   13
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
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
         Left            =   9720
         TabIndex        =   23
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Delete"
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
         Left            =   8280
         TabIndex        =   22
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Active"
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
         Left            =   6600
         TabIndex        =   21
         Top             =   7080
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Edit"
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
         Left            =   4920
         TabIndex        =   20
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "View"
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
         Left            =   840
         TabIndex        =   19
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   720
         TabIndex        =   6
         Top             =   2280
         Width           =   10455
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Gender"
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
            TabIndex        =   12
            Top             =   3720
            Width           =   1695
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Member Type"
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
            Width           =   1455
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Phone Number"
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
            Width           =   1695
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Permanenet Address"
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
            TabIndex        =   9
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Local Address"
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
            TabIndex        =   8
            Top             =   960
            Width           =   1815
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
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   720
         TabIndex        =   2
         Top             =   1200
         Width           =   10455
         Begin VB.CommandButton Command1 
            Caption         =   "Search"
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
            Left            =   8280
            TabIndex        =   5
            Top             =   240
            Width           =   1455
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
            Left            =   3960
            TabIndex        =   4
            Text            =   "(Select)"
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Member ID"
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
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Member Details"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   4560
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Member_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If rs.State = 1 Then rs.Close
    SQL = "select * from member1 where member_id='" + Combo1.Text + "' "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic

    mem_name.Text = (rs.Fields(1))
    mem_local.Text = (rs.Fields(2))
    mem_per.Text = (rs.Fields(3))
    mem_phone.Text = (rs.Fields(4))
    mem_type.Text = (rs.Fields(5))
    mem_gender.Text = (rs.Fields(6))
      
    
End Sub

Private Sub Command2_Click()
member_full_details.Show
End Sub

Private Sub Command3_Click()
 If rs.State = 1 Then rs.Close
    SQL = "update member1 set mem_name='" + mem_name.Text + "',mem_local='" + mem_local.Text + "',mem_per='" + mem_per.Text + "',mem_phone='" + mem_phone.Text + "',mem_type='" + mem_type.Text + "',mem_gender='" + mem_gender.Text + "' where member_id='" + Combo1.Text + "' "
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
    SQL = "delete from member1 where member_id='" + Combo1.Text + "'"
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    Else
    MsgBox "Deleted...!", vbInformation
    End If
End Sub

Private Sub Command6_Click()
Unload Me

End Sub

Private Sub Form_Load()
 If rs.State = 1 Then rs.Close
    SQL = "select * from member1"
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    While Not rs.EOF
        Combo1.AddItem (rs.Fields(0))
        rs.MoveNext
    Wend
    rs.Close
    Command5.Enabled = False
End Sub


Private Sub mem_gender_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
    Case vbKey0 To vbKey9
    KeyAscii = 0
    Beep
    MsgBox "Numeric Not Allowed", vbCritical
    End Select
End Sub

Private Sub mem_name_KeyPress(KeyAscii As Integer)
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

Private Sub mem_phone_KeyUp(KeyCode As Integer, Shift As Integer)
    Call numerictest(mem_phone)
End Sub


Private Sub mem_type_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
    Case vbKey0 To vbKey9
    KeyAscii = 0
    Beep
    MsgBox "Numeric Not Allowed", vbCritical
    End Select
End Sub
