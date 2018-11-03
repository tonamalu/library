VERSION 5.00
Begin VB.Form changepass 
   Caption         =   "Change password"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8115
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   ScaleHeight     =   5355
   ScaleWidth      =   8115
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3660
         TabIndex        =   1
         Top             =   3720
         Width           =   2475
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   6
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   2655
      End
   End
End
Attribute VB_Name = "changepass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text <> "" And Text2.Text <> "" Then

If rs.State = 1 Then rs.Close
    SQL = "select * from login where password='" + Text1.Text + "'"
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic


If Not rs.EOF Then

If rs.State = 1 Then rs.Close
SQL = "update login set password='" & Text2.Text & "' where user_name='admin'"
rs.Open SQL, con, adOpenDynamic, adLockOptimistic



MsgBox "successfully changed password", vbInformation
MsgBox "Plese Login agine", vbInformation
End
Else
MsgBox "wrong username & password", vbExclamation
End If
End If
End Sub

Private Sub Form_Load()
Call connect

End Sub

