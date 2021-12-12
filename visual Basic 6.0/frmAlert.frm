VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notifications"
   ClientHeight    =   8700
   ClientLeft      =   4575
   ClientTop       =   1695
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   15270
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3630
      ItemData        =   "frmAlert.frx":0000
      Left            =   600
      List            =   "frmAlert.frx":0002
      TabIndex        =   2
      Top             =   4800
      Width           =   14175
   End
   Begin VB.ListBox list1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3630
      ItemData        =   "frmAlert.frx":0004
      Left            =   600
      List            =   "frmAlert.frx":0006
      TabIndex        =   0
      Top             =   1200
      Width           =   14175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".... To Do List ...."
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   4740
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Module1.Connect
 Module1.retData "Select * from stock where StockLevel<=10"
With rs
 Do Until .EOF
    list1.AddItem "Buy Stock of " & rs!ProductName
  .MoveNext
  Loop
End With

Module1.retData "Select * from Dealer where outstandings>0"
With rs
 Do Until .EOF
    List2.AddItem "Outstandings Payable to -  " & rs!DealerName
  .MoveNext
  Loop
End With

End Sub

