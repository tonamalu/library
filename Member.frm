VERSION 5.00
Begin VB.Form Member 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Details"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9090
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
      Left            =   4800
      TabIndex        =   13
      Top             =   5880
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
      Left            =   4800
      TabIndex        =   12
      Top             =   5160
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
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   11
      Top             =   4485
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
      Height          =   615
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3720
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
      Height          =   735
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
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
         Left            =   5640
         TabIndex        =   16
         Top             =   7080
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
         Left            =   4200
         TabIndex        =   15
         Top             =   7080
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
         Left            =   2400
         TabIndex        =   14
         Top             =   7080
         Width           =   1455
      End
      Begin VB.TextBox mem_name 
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
         Height          =   405
         Left            =   4800
         TabIndex        =   8
         Top             =   2160
         Width           =   3015
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
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   6120
         Width           =   1455
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
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   5400
         Width           =   1815
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
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   4560
         Width           =   1935
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
         TabIndex        =   4
         Top             =   3840
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
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   3000
         Width           =   1695
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
         TabIndex        =   2
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Member Details"
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
         Left            =   3000
         TabIndex        =   1
         Top             =   720
         Width           =   2655
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

