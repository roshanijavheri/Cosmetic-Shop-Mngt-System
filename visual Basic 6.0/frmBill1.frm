VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBill 
   Caption         =   "Bill"
   ClientHeight    =   8595
   ClientLeft      =   4695
   ClientTop       =   855
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   13770
   Begin VB.CommandButton cmdHistory 
      Caption         =   "History"
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
      Left            =   11760
      TabIndex        =   24
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   7320
      TabIndex        =   20
      Tag             =   "For Your Help"
      ToolTipText     =   "Enter Quantity"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      DataField       =   "ProductName"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Index           =   1
      Left            =   1080
      TabIndex        =   14
      Top             =   3600
      Width           =   5895
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Index           =   2
      Left            =   6960
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Index           =   3
      ItemData        =   "frmBill1.frx":0000
      Left            =   8520
      List            =   "frmBill1.frx":0002
      TabIndex        =   11
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4230
      Index           =   4
      Left            =   9960
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdBill 
      Caption         =   "Get Bill"
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
      Left            =   9480
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
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
      Left            =   11760
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox txtTotalBill 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9720
      TabIndex        =   6
      Text            =   "0"
      Top             =   7920
      Width           =   1815
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
      Left            =   4800
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtCustomer 
      BorderStyle     =   0  'None
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   6255
   End
   Begin VB.TextBox txtBillno 
      BorderStyle     =   0  'None
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
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1335
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
      ItemData        =   "frmBill1.frx":0004
      Left            =   240
      List            =   "frmBill1.frx":0006
      TabIndex        =   0
      Text            =   "Select Products"
      Top             =   2160
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   615
      Left            =   8760
      TabIndex        =   23
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      _Version        =   393216
      MousePointer    =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Thank You !!! Visit Again !!!"
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
      Top             =   8160
      Width           =   4320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   9960
      TabIndex        =   21
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   630
      Left            =   4920
      TabIndex        =   19
      Top             =   0
      Width           =   1740
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sr."
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Products"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1080
      TabIndex        =   17
      Top             =   3120
      Width           =   5895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   8520
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6960
      TabIndex        =   15
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Total Bill ="
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   7
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No :"
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
      TabIndex        =   4
      Top             =   600
      Width           =   1230
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill To :"
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
      TabIndex        =   3
      Top             =   1440
      Width           =   1155
   End
End
Attribute VB_Name = "frmBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, CurrentAmount, AvailableStock As Integer
Dim rs1 As New Recordset     ' For accessing Max BillID Value From Bill
Dim rs2 As New Recordset    ' For BillConvinience
Dim rs3 As New Recordset   'for Product Table


Private Sub cmdBill_Click()
   Dim i, Stock As Integer
    Module1.retData "Select * from Bill"
    rs.AddNew
    For i = 0 To list1(1).ListCount - 1
        rs.Fields("BillNo") = txtBillno.Text
        rs.Fields("Date") = Date1.Value
        rs.Fields("CustomerName") = txtCustomer.Text
        rs.Fields("ProductName") = list1(1).List(i)
        rs.Fields("Qty") = list1(3).List(i)
        rs.Fields("TotalBill") = "0"
        rs.Update
        rs.AddNew
    Next
    rs.Fields("CustomerName") = txtCustomer.Text
    rs.Fields("TotalBill") = txtTotalBill.Text
    rs.Update
    MsgBox "Bill Has Been Saved Successfully!!!!"
                                                                ' Bill Has been Saved in Bill Table
    
    cmdBill.Visible = False
    cmdRemove.Visible = False
    Text1.Visible = False
    Combo1.Visible = False
    Label2.Top = 2160       'To Move Objects Upside
    Label4.Top = 2160
    Label5.Top = 2160
    Label6.Top = 2160
    Label7.Top = 2160
    list1(0).Top = 2640
    list1(1).Top = 2640
    list1(2).Top = 2640
    list1(3).Top = 2640
    list1(4).Top = 2640
    txtTotalBill.Top = 6880
    Label3.Top = 6880
     
    rs2.Open "Select * from BillConvinience", Con, adOpenStatic, adLockOptimistic
    rs2.AddNew
    rs2!BillNo = txtBillno.Text
    rs2!CustomerName = txtCustomer.Text
    rs2!BillAmount = CurrentAmount
    rs2!Date = Date1.Value
    rs2.Update
                                      ' Details Has been Saved in bill Convinience
                                      
       'Dont Remove this Code VIMP Code in this Form For Updation in Products Qty
     Dim Item As String
     For i = 0 To list1(1).ListCount - 1
     Item = list1(1).List(i)
     rs3.Open "Select * from Product Where ProductName = '" & Item & "' ", Con, adOpenStatic, adLockOptimistic
     Stock = rs3.Fields("StockLevel")
     rs3.Fields("StockLevel") = Stock - list1(3).List(i)
     rs3.Update
     rs3.Close
     Next i
        'VIMP Till Here
    
    
    Me.PrintForm
    Call Clear
    rs.AddNew
End Sub

Private Sub Clear()
txtBillno.Text = ""
txtCustomer.Text = ""
Text3.Text = ""
Text1.Text = ""
Combo1.Text = "Select Product"
list1.Item(0).Clear
list1.Item(1).Clear
list1.Item(2).Clear
list1.Item(3).Clear
list1.Item(4).Clear
txtTotalBill.Text = "0"
End Sub

Private Sub cmdHistory_Click()
frmBillHistory.Show
End Sub

Private Sub cmdRemove_Click()
list1(2).RemoveItem list1(1).ListIndex
CurrentAmount = CurrentAmount - (list1(4).List(list1(1).ListIndex))   ' Subtract Amount of Product Removed
txtTotalBill.Text = CurrentAmount                                     ' For Displaying The TotalAmount
list1(3).RemoveItem list1(1).ListIndex
list1(4).RemoveItem list1(1).ListIndex
list1(1).RemoveItem list1(1).ListIndex
list1(0).Clear
For itr = 0 To list1(1).ListCount - 1
  list1(0).AddItem itr + 1
Next
i = i - 1
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Dim Amount1, NewQty As Integer
If Combo1.Text <> "" Then
Module1.retData "Select * from Product Where ProductName='" & Combo1 & "' "
'Don't Remove This Code. This is For Duplicate Product Entries{
 For itr = 0 To list1(1).ListCount - 1
    If (Combo1.Text = list1(1).List(itr)) Then
      NewQty = InputBox("You Have Selected " & Combo1.Text & "again" & vbNewLine & "Enter New Qty", "Same Product Selected!!!")
      If (NewQty <= 0) Then
       Exit Sub
      Else
      list1(3).List(itr) = NewQty
      list1(4).List(itr) = NewQty * list1(2).List(itr)
       For itr1 = 0 To list1(4).ListCount - 1
        Amount1 = Amount1 + Val(list1(4).List(itr1))
        CurrentAmount = Amount1
        txtTotalBill.Text = CurrentAmount
       Next
      Combo1.SetFocus
      Exit Sub
      End If
      End If
 Next itr
 
'Dont Remove - VIMP Code}
 
list1(0).AddItem i                                   'Code For Sr No
i = i + 1
list1(1).AddItem Combo1.Text                               ' Code For Product Name In List Box
list1(2).AddItem rs.Fields("SellingPrice")             ' Code For Amount Section
Text3.Text = rs.Fields("SellingPrice")                    'this is useful for Converting Selling Price in Integer Format
AvailableStock = rs.Fields("StockLevel")
Text1.SetFocus
End If
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
Module1.retData "Select ProductName from Product where ProductName Like '" & Combo1.Text & "'  "
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
i = 1
Module1.Connect
Module1.retData "Select * from Product"
With rs
 Do Until .EOF
 Combo1.AddItem ![ProductName]
 .MoveNext
 Loop
End With
rs1.Open ("Select Max(BillNo) from Bill"), Con, adOpenStatic, adLockOptimistic
txtBillno.Text = rs1.Fields(0).Value + 1
Date1.Value = Date
End Sub

Private Sub Label11_Click()
txtCustomer.SetFocus
End Sub

Private Sub Label17_Click()
txtBillno.SetFocus
End Sub

Private Sub List1_Click(Index As Integer)      ' For Selecting Whole Row at a time
On Error GoTo Trap
    list1(0).ListIndex = list1(1).ListIndex
    list1(2).ListIndex = list1(1).ListIndex
    list1(3).ListIndex = list1(1).ListIndex
    list1(4).ListIndex = list1(1).ListIndex
Trap:
 Exit Sub
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

Dim Calculate As Integer
If (list1(1).ListCount > list1(3).ListCount) Then
 If KeyAscii = 13 Then
  On Error GoTo Trap
     If (Val(Text1.Text) <= 0) Then
      MsgBox "You Have Entered" & Text1.Text, vbCritical + vbExclamation, "Cosmeta"
      Exit Sub
     End If
      If (AvailableStock < Text1.Text) Then
      MsgBox "Available Only " & AvailableStock, vbCritical + vbExclamation, "Cosmeta"
      Exit Sub
     End If
    list1(3).AddItem (Text1.Text)
    list1(4).AddItem (Val(Text3.Text) * Val(Text1.Text))
    Calculate = (Val(Text3.Text) * Val(Text1.Text))        ' For Converting The List Item Into Integer
    CurrentAmount = CurrentAmount + Calculate         ' Change this if the Item is Deleted
    txtTotalBill.Text = CurrentAmount
    Text1.Text = ""
    Combo1.SetFocus
 End If
End If
Trap:
  Exit Sub
 
End Sub



Private Sub txtBillno_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  txtCustomer.SetFocus
 End If
End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Combo1.SetFocus
End If
End Sub

