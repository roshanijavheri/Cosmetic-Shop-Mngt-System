VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13890
   BeginProperty Font 
      Name            =   "Nirmala UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   13890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000B&
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   12120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Show Password"
      Top             =   4440
      Value           =   2  'Grayed
      Width           =   210
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   3495
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Enter Your Password"
      Top             =   4200
      Width           =   3735
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   8760
      TabIndex        =   2
      ToolTipText     =   "Please Enter Your Username"
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label lblPass 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   735
      Left            =   5640
      TabIndex        =   1
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label lblUsername 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   735
      Left            =   5625
      TabIndex        =   0
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   -120
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   14070
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Global Declaration
Dim user As String
Dim password As String
'To Show & Hide the Password
Private Sub Check1_Click()
If Check1.Value = 1 Then
  Check1.ToolTipText = "Hide Password"
  txtPass.PasswordChar = ""
Else
  Check1.ToolTipText = "Show Password"
  txtPass.PasswordChar = "*"
End If
End Sub

Private Sub cmdLogin_Click()
user = "Cosmeta"
password = "123"

If (user = txtUsername.Text And password = txtPass.Text) Then
    HomePage.Show
    MsgBox "Welcome To Cosmeta", , "Cosmeta"
    Unload Me
Else
    MsgBox "Username or Password is Not Correct"
End If

End Sub



Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdLogin_Click
End If

End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPass.SetFocus
End If
End Sub
