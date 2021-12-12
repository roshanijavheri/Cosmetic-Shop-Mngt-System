VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSearchRepo 
   Caption         =   "Search Reports"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   15195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   11
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Show All"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   10
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   8
      Top             =   3840
      Width           =   2055
   End
   Begin VB.OptionButton optStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Staff"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10680
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton optDealer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dealer"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton optBill 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bill"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.OptionButton optProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Products"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker Date1 
      Height          =   735
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      _Version        =   393216
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   138412033
      CurrentDate     =   44247
   End
   Begin MSComCtl2.DTPicker Date2 
      Height          =   735
      Left            =   8280
      TabIndex        =   1
      Top             =   2160
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      _Version        =   393216
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   138412033
      CurrentDate     =   44247
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Price to Search Products Under Price :"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   9
      Top             =   3840
      Width           =   8055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   2205
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2235
      Width           =   975
   End
End
Attribute VB_Name = "FrmSearchRepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAll_Click()
    If (optBill.Value = True) Then
        Module1.retData "Select * from BillConvinience"
        
        If DataEnvironment1.rsBillConvinience.State = adStateOpen Then
         DataEnvironment1.rsBillConvinience.Close
        End If
        
        DataEnvironment1.BillConvinience CDate("01/01/1000"), CDate("01/01/2100")
        Load BillConvinienceReport1
        BillConvinienceReport1.Show
    End If
End Sub

Private Sub cmdPrint_Click()
 If (optBill.Value = True) Then
    If DataEnvironment1.rsBillConvinience.State = adStateOpen Then
     DataEnvironment1.rsBillConvinience.Close
    End If
    
    DataEnvironment1.BillConvinience Date1.Value, Date2.Value
    Load BillConvinienceReport1
    BillConvinienceReport1.Show
 End If
End Sub

Private Sub Form_Load()
Module1.Connect
End Sub

