VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProducts 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Products"
   ClientHeight    =   10545
   ClientLeft      =   4620
   ClientTop       =   780
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   15855
   Begin VB.CommandButton cmdBack 
      Caption         =   "<-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13200
      TabIndex        =   16
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   13
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox txtCostPrice 
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2640
      TabIndex        =   5
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      TabIndex        =   14
      Top             =   8400
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      TabIndex        =   15
      Top             =   8400
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "@Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      ItemData        =   "frmProducts.frx":0000
      Left            =   2640
      List            =   "frmProducts.frx":0052
      TabIndex        =   4
      Top             =   3360
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7335
      Left            =   6600
      TabIndex        =   17
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12938
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   30
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
         Size            =   14.25
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
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   10
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   12
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   11
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   8400
      Width           =   1935
   End
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   7200
      Width           =   3375
   End
   Begin VB.TextBox txtProductNo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtProductName 
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtStock 
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2640
      TabIndex        =   7
      Top             =   6240
      Width           =   3375
   End
   Begin VB.TextBox txtPrice 
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2640
      TabIndex        =   6
      Top             =   5280
      Width           =   3375
   End
   Begin VB.TextBox txtBrand 
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5640
      Picture         =   "frmProducts.frx":0175
      Stretch         =   -1  'True
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Price"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   360
      TabIndex        =   24
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   240
      TabIndex        =   23
      Top             =   7200
      Width           =   1830
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Price"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   240
      TabIndex        =   22
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   2310
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   600
      TabIndex        =   20
      Top             =   2400
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   360
      TabIndex        =   19
      Top             =   3360
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   600
      TabIndex        =   18
      Top             =   6240
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product No."
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuAllProduct 
         Caption         =   "All Product's Report"
      End
      Begin VB.Menu mnuAmountWise 
         Caption         =   "Amount Wise Report"
      End
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New Recordset

Private Sub cmdAddNew_Click()
On Error GoTo Trap
    Call cmdClear_Click
    rs.AddNew
    rs1.Open ("Select (Max(ProductID)) from Product"), Con, adOpenStatic, adLockOptimistic
    txtProductNo.Text = rs1.Fields(0).Value + 1
    rs1.Close
Trap:
 Exit Sub
End Sub

Private Sub cmdBack_Click()
On Error GoTo Trap
Module1.retData "Select * from Product"
Call LoadData
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
cmdBack.Visible = False
Trap:
 Exit Sub
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Trap
Dim Confirm As String
Confirm = MsgBox("Are You Sure, You Want To delete this " & rs!ProductName & " Record ?", vbYesNo + vbCritical, "Delete Record Confirmation!!!")
If Confirm = vbYes Then
 rs.Delete
 MsgBox "Record deleted Successfully", vbInformation, "Cosmeta"
 rs.Update
 Call LoadData
 rs.MoveNext
End If
Trap:
 Exit Sub
End Sub

Private Sub cmdLast_Click()
On Error GoTo Trap
rs.MoveLast
Call LoadData
Trap:
 Exit Sub
End Sub

Private Sub cmdNext_Click()
On Error GoTo Trap
rs.MoveNext
If (rs.EOF = True) Then
 rs.MoveFirst
End If
Call LoadData
Trap:
 Exit Sub
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo Trap
rs.MovePrevious
If (rs.BOF = True) Then
 rs.MoveLast
End If
Call LoadData
Trap:
 Exit Sub
End Sub



Private Sub cmdSave_Click()
On Error Resume Next
    If (txtProductNo.Text = "") And (txtProductName.Text = "") And (txtBrand.Text = "") And (Combo1.Text = "") And (txtCostPrice.Text = "") And (txtPrice.Text = "") And (txtStock.Text = "") Then
        MsgBox "Please Enter Some Data", vbInformation, "Cosmeta"
    Else
        rs.Fields(0) = txtProductNo.Text      'Dont uncomment It Because ID is Not Updatable You Have Assigned PID as Auto Number
        rs.Fields("ProductName") = txtProductName.Text
        rs.Fields(2) = txtBrand.Text
        rs.Fields("Category") = Combo1.Text
        rs!CostPrice = txtCostPrice
        rs!SellingPrice = txtPrice.Text
        rs!StockLevel = txtStock.Text
        rs!Description = txtDescription.Text
    
        rs.Update
        MsgBox ("Record Saved!!!!")
    
    End If
    
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
 Module1.Connect
 Module1.retData "Select * from Product"
 Call LoadData
 Set DataGrid1.DataSource = rs
 DataGrid1.Refresh
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
 PopupMenu mnuPopup, vbPopupMenuRightButton
End If
End Sub
Private Sub LoadData()

    txtProductNo.Text = rs.Fields(0)
    txtProductName.Text = rs.Fields("ProductName")
    txtBrand.Text = rs.Fields(2)
    Combo1.Text = rs.Fields("Category")
    txtCostPrice = rs!CostPrice
    txtPrice.Text = rs!SellingPrice
    txtStock.Text = rs!StockLevel
    txtDescription.Text = rs!Description
    
End Sub


Private Sub cmdClear_Click()
    txtProductNo.Text = ""
    txtProductName.Text = ""
    txtBrand.Text = ""
    Combo1.Text = ""
    txtCostPrice.Text = ""
    txtPrice.Text = ""
    txtStock.Text = ""
    txtDescription.Text = ""
End Sub



Private Sub cmdFirst_Click()
On Error GoTo Trap
rs.MoveFirst
Call LoadData
Trap:
 Exit Sub
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCostPrice.SetFocus
End If
End Sub


Private Sub Image1_Click()
On Error GoTo Trap
Module1.retData "Select * from Product Where ProductName like '%" & txtProductName.Text & "%' "
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Call LoadData
  cmdBack.Visible = True
Trap:
 Exit Sub
End Sub

Private Sub mnuAllProduct_Click()     ' Static Report
ProductReport.Show
End Sub

Private Sub mnuAmountWise_Click()             'For Dynamic Report
If (txtPrice.Text = "") Then
 MsgBox "Enter Selling Price to search", vbInformation, "Cosmeta"
Else
    DataEnvironment1.ProductAmountWise txtPrice.Text
    Load ProductReportAmountWise
    ProductReportAmountWise.Show
    ProductReportAmountWise.Refresh
    DataEnvironment1.rsProductAmountWise.Close
End If
End Sub

Private Sub txtBrand_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
End Sub

Private Sub txtCostPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPrice.SetFocus
End If
End Sub

Private Sub txtCostPrice_KeyUp(KeyCode As Integer, Shift As Integer)
If (Not IsNumeric(txtCostPrice.Text)) Then
 MsgBox "Enter Valid Price", vbExclamation, "Wrong Value Entered!!!"
 txtCostPrice.Text = ""
End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtStock.SetFocus
End If
End Sub

Private Sub txtPrice_KeyUp(KeyCode As Integer, Shift As Integer)
If (Not IsNumeric(txtPrice.Text)) Then
 MsgBox "Enter Valid Price", vbExclamation, "Wrong Value Entered!!!"
 txtPrice.Text = ""
End If
End Sub

Private Sub txtProductName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBrand.SetFocus
End If
End Sub

Private Sub txtProductNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtProductName.SetFocus
End If
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDescription.SetFocus
End If
End Sub

Private Sub txtStock_KeyUp(KeyCode As Integer, Shift As Integer)
If (Not IsNumeric(txtStock.Text)) Then
 MsgBox "Enter Valid Price", vbExclamation, "Wrong Value Entered!!!"
 txtStock.Text = ""
End If
End Sub
