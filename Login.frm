VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.TextBox TextBox2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Left            =   1620
         TabIndex        =   2
         Top             =   1260
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2940
         TabIndex        =   5
         Top             =   1980
         Width           =   1395
      End
      Begin VB.TextBox TextBox1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Left            =   1620
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call login
End Sub

Private Sub TextBox1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        TextBox2.SetFocus
        KeyAscii = 0
    End If
End Sub
Private Sub TextBox2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        Call login
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox2_Change()
    TextBox2.PasswordChar = "*"
End Sub

Private Sub login()
    connect
    Dim rstuser As New ADODB.Recordset
    'Dim loginok As Boolean
    'Debug.Print con.State
    If TextBox1.Text <> "" And TextBox2.Text <> "" Then
    rstuser.Open "select * from login where user_name='" + TextBox1.Text + "'" + "and password='" + TextBox2.Text + "'", con, adOpenDynamic, adLockOptimistic
    If Not rstuser.EOF Then
    Unload Me
    library.Show
    Else
    MsgBox "wrong username & password", vbExclamation
    End If
    End If
End Sub

