VERSION 5.00
Begin VB.Form SplashScreen 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   70
      Left            =   360
      Top             =   360
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   555
      Left            =   6240
      TabIndex        =   3
      Top             =   4440
      Width           =   1650
   End
   Begin VB.Label lblProjectBy 
      BackStyle       =   0  'Transparent
      Caption         =   "Project By:                          34. Roshani M. Javheri          36. Aishwarya V. Kalamani"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   1620
      Left            =   9600
      TabIndex        =   2
      Top             =   6360
      Width           =   4395
   End
   Begin VB.Label Percentage 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Shape shpProgress 
      BackColor       =   &H00C0E0FF&
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   75
   End
   Begin VB.Shape shpProgressBar 
      BorderColor     =   &H00C0E0FF&
      Height          =   495
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   7935
   End
   Begin VB.Label lblCosmeta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cosmeta"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   126
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2910
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   10860
   End
   Begin VB.Image imgSS 
      Height          =   8115
      Left            =   -120
      Picture         =   "SplashScreen.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   14280
   End
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr, ctr2, r As Double
Dim ctr3 As String

Private Sub Timer1_Timer()
ctr = 0
If ctr2 <= 100 Then
    Randomize
    r = Int((200 - 100 + 1) * Rnd + 100)
    ctr = r / 50
    ctr = Round(ctr, 0)
    ctr = ctr2 + ctr
    ctr3 = Str(ctr)
     If ctr >= 100 Then
        Percentage.Caption = "100%"
        shpProgress.Width = 7935
        Login.Show
        Unload Me
    Else
        shpProgress.Width = shpProgress.Width + r
        'Percentage.Caption = (ctr3) + "%"
        ctr2 = Int(ctr3)
    End If
End If

End Sub
