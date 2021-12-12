VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAllPO 
   Caption         =   "All Purchase Order"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Enter Product Name"
      Top             =   720
      Width           =   9375
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
      Left            =   11160
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6975
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   12303
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   37
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Microsoft YaHei"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Microsoft YaHei UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAllPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Module1.Connect
Module1.retData "Select * from PurchaseOrder"
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Module1.retData "Select * from PurchaseOrder Where DealerName Like '%" & Text1.Text & "%' "
 Set DataGrid1.DataSource = rs
 DataGrid1.Refresh
End Sub


Private Sub cmdAll_Click()
Module1.retData "Select * from PurchaseOrder"
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub



