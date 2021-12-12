VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBillHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Histoty"
   ClientHeight    =   9330
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5880
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Enter Amount to search"
      Top             =   1080
      Width           =   1575
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
      Left            =   7680
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
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
      Height          =   705
      Left            =   240
      TabIndex        =   2
      Text            =   "Enter Customer Name"
      ToolTipText     =   "Enter Customer Name"
      Top             =   1080
      Width           =   5295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6975
      Left            =   240
      TabIndex        =   1
      Top             =   2040
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
   Begin MSComCtl2.DTPicker Date1 
      Height          =   735
      Left            =   10080
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      _Version        =   393216
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   138477571
      CurrentDate     =   44247
   End
   Begin MSComCtl2.DTPicker Date2 
      Height          =   735
      Left            =   12480
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      _Version        =   393216
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   138477571
      CurrentDate     =   44275
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPrintAll 
         Caption         =   "Print All"
      End
      Begin VB.Menu mnuPrintAmoutwise 
         Caption         =   "Print Amount Wise"
      End
      Begin VB.Menu mnuPrintDateWise 
         Caption         =   "Print Between Dates"
      End
   End
End
Attribute VB_Name = "frmBillHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Date1_Change()
 Module1.retData "Select * from BillConvinience Where date Like '" & Date1 & "' "
 Set DataGrid1.DataSource = rs
 DataGrid1.Refresh
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Module1.Connect
Module1.retData "Select * from BillConvinience"
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Date1.Value = Date
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
 PopupMenu mnuPopup, vbPopupMenuRightButton
End If
End Sub

Private Sub mnuPrintAll_Click()
    Load BillConvinienceReport
    BillConvinienceReport.Show
    BillConvinienceReport.Refresh
    DataEnvironment1.rsBillConvinience.Close
End Sub

Private Sub mnuPrintAmoutwise_Click()
If (Text2.Text = "") Then
 MsgBox "Enter Amount To search Bills", vbInformation, "Cosmeta"
Else
    DataEnvironment1.Bill1 Text2.Text
    Load BillAboveDataReport
    BillAboveDataReport.Show
    BillAboveDataReport.Refresh
    DataEnvironment1.rsBill1.Close
End If
End Sub

Private Sub mnuPrintDateWise_Click()
DataEnvironment1.BillDateWise Date1.Value, Date2.Value
Load BillDatewiseReport
BillDatewiseReport.Show
BillDatewiseReport.Refresh
DataEnvironment1.rsBillDateWise.Close
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Module1.retData "Select * from BillConvinience Where CustomerName Like '%" & Text1.Text & "%' "
 Set DataGrid1.DataSource = rs
 DataGrid1.Refresh
End Sub


Private Sub cmdAll_Click()
Module1.retData "Select * from BillConvinience"
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
End Sub


Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Module1.retData "Select * from BillConvinience Where BillAmount Like '%" & Text2.Text & "%'"
 Set DataGrid1.DataSource = rs
 DataGrid1.Refresh
End Sub
