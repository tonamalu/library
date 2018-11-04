VERSION 5.00
Begin VB.Form Member 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Details"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8130
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
      Left            =   4860
      TabIndex        =   11
      Top             =   5340
      Width           =   3015
   End
   Begin VB.TextBox mem_type 
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
      Left            =   4860
      TabIndex        =   10
      Top             =   4680
      Width           =   3015
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
      Height          =   450
      Left            =   4860
      MaxLength       =   10
      TabIndex        =   9
      Top             =   4065
      Width           =   3015
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
      Height          =   915
      Left            =   4860
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   3015
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
      Height          =   795
      Left            =   4860
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Member Details "
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   -60
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.TextBox mem_name 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   4920
         TabIndex        =   15
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CommandButton exit 
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
         Left            =   6600
         TabIndex        =   14
         Top             =   6300
         Width           =   1335
      End
      Begin VB.CommandButton clear 
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
         Left            =   5280
         TabIndex        =   13
         Top             =   6300
         Width           =   1215
      End
      Begin VB.CommandButton add 
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
         Left            =   3720
         TabIndex        =   12
         Top             =   6300
         Width           =   1455
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   360
         TabIndex        =   6
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   360
         TabIndex        =   4
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Permanent Address"
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
         Left            =   360
         TabIndex        =   3
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
 If rs.State = 1 Then rs.Close
 If mem_name.Text = "" Then
 MsgBox ("fill the Member Name fields"), vbInformation
 ElseIf mem_local.Text = "" Then
 
 MsgBox ("fill  the Local address"), vbInformation
  ElseIf mem_per.Text = "" Then
 
 MsgBox ("fill  the Permenant address"), vbInformation

   ElseIf mem_phone.Text = "" Then

 MsgBox ("fill  the Phone number "), vbInformation
   ElseIf mem_type.Text = "" Then
 
 MsgBox ("fill  the member type "), vbInformation

    ElseIf mem_gender.Text = "" Then
 
 MsgBox ("fill  the gender "), vbInformation

    ElseIf mem_phone.Text < 7000000000# Then
 
 MsgBox ("enter a valid phone number  "), vbInformation

 Else
 
 
    SQL = "insert into member1(mem_name,mem_local,mem_per,mem_phone,mem_type,mem_gender)values('" + mem_name.Text + "','" + mem_local.Text + "','" + mem_per.Text + "','" + mem_phone.Text + "','" + mem_type.Text + "','" + mem_gender.Text + "') "
    rs.Open SQL, con, adOpenDynamic, adLockOptimistic
    MsgBox ("succesfully completed"), vbInformation
   
    End If
    
End Sub

Private Sub clear_Click()
'mem_id.Text = " "
mem_name.Text = " "
mem_local.Text = " "
mem_per.Text = " "
mem_phone.Text = " "
mem_type.Text = " "
mem_gender.Text = " "

End Sub

Private Sub exit_Click()
Unload Me
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

