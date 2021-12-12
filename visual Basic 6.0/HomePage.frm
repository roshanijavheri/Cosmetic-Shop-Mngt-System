VERSION 5.00
Begin VB.Form HomePage 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cosmeta (Homepage)"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   11415
      Left            =   0
      MouseIcon       =   "HomePage.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   -120
      Width           =   4575
      Begin VB.Label lblAbout 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Left            =   1800
         TabIndex        =   11
         Top             =   8280
         Width           =   1350
      End
      Begin VB.Image Image8 
         Height          =   615
         Left            =   720
         Picture         =   "HomePage.frx":030A
         Stretch         =   -1  'True
         Top             =   8280
         Width           =   735
      End
      Begin VB.Image Image7 
         Height          =   615
         Left            =   720
         Picture         =   "HomePage.frx":AC3F
         Stretch         =   -1  'True
         Top             =   7320
         Width           =   735
      End
      Begin VB.Label lblAlerts 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alerts"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Left            =   1815
         TabIndex        =   10
         Top             =   7320
         Width           =   1320
      End
      Begin VB.Label lblStock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Left            =   1830
         TabIndex        =   9
         Top             =   6360
         Width           =   1290
      End
      Begin VB.Image Image6 
         Height          =   615
         Left            =   720
         Picture         =   "HomePage.frx":11EED
         Stretch         =   -1  'True
         Top             =   6360
         Width           =   735
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   720
         Picture         =   "HomePage.frx":1B388
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   720
         Picture         =   "HomePage.frx":21036A
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   735
      End
      Begin VB.Image Image5 
         Height          =   615
         Left            =   720
         Picture         =   "HomePage.frx":21AEAD
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   735
      End
      Begin VB.Image Image4 
         Height          =   735
         Left            =   720
         Picture         =   "HomePage.frx":2474D3
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   720
         Picture         =   "HomePage.frx":24A460
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label lblBill 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Left            =   1995
         TabIndex        =   8
         Top             =   2520
         Width           =   690
      End
      Begin VB.Label lblStaff 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Left            =   1845
         TabIndex        =   7
         Top             =   3480
         Width           =   1080
      End
      Begin VB.Label lblOrder 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Left            =   1815
         TabIndex        =   6
         Top             =   4440
         Width           =   1320
      End
      Begin VB.Label lblDealer 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealers"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Left            =   1665
         TabIndex        =   5
         Top             =   5400
         Width           =   1785
      End
      Begin VB.Label lblProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Products"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Left            =   1650
         TabIndex        =   3
         Top             =   1680
         Width           =   2040
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cosmeta"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4335
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   1335
         Left            =   0
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   11535
      Left            =   4560
      TabIndex        =   0
      Top             =   -240
      Width           =   15855
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   825
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   2280
      End
      Begin VB.Image Image9 
         Height          =   11010
         Left            =   0
         Picture         =   "HomePage.frx":253A98
         Stretch         =   -1  'True
         Top             =   240
         Width           =   15735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Products"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   540
         Left            =   -1920
         TabIndex        =   4
         Top             =   2640
         Width           =   1740
      End
   End
End
Attribute VB_Name = "HomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub lblAbout_Click()
frmAbout.Show
End Sub

Private Sub lblAlerts_Click()
frmAlert.Show
End Sub

Private Sub lblBill_Click()
frmBill.Show
End Sub

Private Sub lblDealer_Click()
frmDealer.Show
End Sub

Private Sub lblOrder_Click()
frmOrder.Show
End Sub

Private Sub lblProduct_Click()
frmProducts.Show
End Sub

Private Sub lblStaff_Click()
frmStaff.Show
End Sub

'Code for Closing Entire software on Exit Button
Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lblStock_Click()
frmStock.Show
End Sub

