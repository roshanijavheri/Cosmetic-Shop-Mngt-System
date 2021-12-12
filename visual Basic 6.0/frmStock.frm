VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock"
   ClientHeight    =   8700
   ClientLeft      =   4575
   ClientTop       =   1695
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Microsoft YaHei"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   15270
   Begin VB.CommandButton cmdLow 
      Caption         =   "Low Stock"
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
      Left            =   12360
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   705
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Enter Product Name"
      Top             =   600
      Width           =   8295
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
      Left            =   9960
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6975
      Left            =   240
      TabIndex        =   0
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
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAllStock 
         Caption         =   "Available Stock Report"
      End
      Begin VB.Menu mnuLowStock 
         Caption         =   "Low Stock Report"
      End
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAll_Click()
Module1.retData "Select * from stock"
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub

Private Sub cmdLow_Click()
Module1.retData "Select * from stock where StockLevel<=10"
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub


Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = "Enter Product to Search"
Module1.Connect
Module1.retData "Select *  from stock"
DataGrid1.DefColWidth = Width / 2
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
 PopupMenu mnuPopup, vbPopupMenuRightButton
End If
End Sub

Private Sub mnuAllStock_Click()
AvailableStockReport.Show
End Sub

Private Sub mnuLowStock_Click()
    Load StockDataReport
    StockDataReport.Show
    StockDataReport.Refresh
    DataEnvironment1.rsStock.Close
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Module1.retData "Select * from Stock Where ProductName like '%" & Text1.Text & "%'"
 Set DataGrid1.DataSource = rs
 DataGrid1.Refresh
End Sub
