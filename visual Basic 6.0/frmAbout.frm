VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Cosmeata"
   ClientHeight    =   8895
   ClientLeft      =   4620
   ClientTop       =   1170
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   13605
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version - 1.0"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   555
      Left            =   4680
      TabIndex        =   5
      Top             =   6840
      Width           =   4170
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "©"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   825
      Left            =   120
      TabIndex        =   4
      Top             =   8160
      Width           =   600
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RG Softwares Pvt. Ltd.  All Rights reserved."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   555
      Left            =   840
      TabIndex        =   3
      Top             =   8280
      Width           =   10050
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3870
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   12330
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "__________"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2310
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   10050
   End
   Begin VB.Label lblCosmeta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cosmeta"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2775
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   8955
   End
   Begin VB.Image Image1 
      Height          =   8955
      Left            =   -120
      Picture         =   "frmAbout.frx":00EA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14040
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Deactivate()
Unload Me
End Sub

