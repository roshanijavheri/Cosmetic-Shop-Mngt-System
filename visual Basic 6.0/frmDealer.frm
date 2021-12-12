VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDealer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dealer's Info"
   ClientHeight    =   8850
   ClientLeft      =   4575
   ClientTop       =   1695
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   15420
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
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7335
      Left            =   6360
      TabIndex        =   23
      Top             =   600
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
      Left            =   120
      TabIndex        =   21
      Top             =   8040
      Width           =   1935
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
      Left            =   4440
      TabIndex        =   20
      Top             =   8040
      Width           =   1935
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
      Left            =   6720
      TabIndex        =   19
      Top             =   8040
      Width           =   1575
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
      Left            =   2400
      TabIndex        =   18
      Top             =   8040
      Width           =   1695
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
      Left            =   12120
      TabIndex        =   17
      Top             =   8040
      Width           =   975
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
      Left            =   10800
      TabIndex        =   16
      Top             =   8040
      Width           =   975
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
      TabIndex        =   15
      Top             =   8040
      Width           =   1695
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
      Left            =   13440
      TabIndex        =   14
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox txtDetails 
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   6480
      Width           =   3375
   End
   Begin VB.TextBox txtDealerContact 
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
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox txtDealerAddress 
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
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox txtEmail 
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
      Left            =   2760
      TabIndex        =   4
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox txtDealerOutstanding 
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
      Left            =   2760
      TabIndex        =   3
      Top             =   5520
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   495
      Left            =   20280
      TabIndex        =   2
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox txtDealerName 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox txtDealerNo 
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
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   5280
      Picture         =   "frmDealer.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Id"
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
      TabIndex        =   22
      Top             =   4560
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
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
      TabIndex        =   13
      Top             =   6600
      Width           =   1080
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer No"
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
      TabIndex        =   11
      Top             =   720
      Width           =   1590
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding"
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
      TabIndex        =   10
      Top             =   5520
      Width           =   1995
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   480
      TabIndex        =   9
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
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
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Name"
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
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAllDealer 
         Caption         =   "All Dealers Reprt"
      End
      Begin VB.Menu mnuDealerOutstanding 
         Caption         =   "Outstanding Pending Dealer's Report "
      End
   End
End
Attribute VB_Name = "frmDealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs1 As New Recordset
Private Sub cmdAddNew_Click()
 Call cmdClear_Click
 rs.AddNew
 cmdSave.SetFocus
 rs1.Open "Select Max(DealerID) from Dealer", Con, adOpenStatic, adLockOptimistic
 txtDealerNo.Text = rs1.Fields(0).Value + 1
 rs1.Close
End Sub

Private Sub cmdBack_Click()
Module1.retData "Select * from Dealer"
Call LoadData
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
cmdBack.Visible = False
End Sub

Private Sub cmdClear_Click()
txtDealerNo.Text = ""
txtDealerName.Text = ""
txtDealerContact.Text = ""
txtDealerAddress.Text = ""
txtEmail.Text = ""
txtDealerOutstanding.Text = ""
txtDetails.Text = ""
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim Confirm As String
Confirm = MsgBox("Are You Sure, You Want To delete this " & rs!DealerName & " Record ?", vbYesNo + vbCritical, "Delete Record Confirmation!!!")
If Confirm = vbYes Then
 rs.Delete
 MsgBox "Record deleted Successfully", vbInformation, "Cosmeta"
 rs.MoveNext
 Call LoadData
 rs.Update
End If
End Sub

Private Sub cmdFirst_Click()
 rs.MoveFirst
 Call LoadData
End Sub

Private Sub cmdLast_Click()
 rs.MoveLast
 Call LoadData
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
 rs.MoveNext
 If rs.EOF Then
  rs.MoveFirst
 End If
 Call LoadData
End Sub

Private Sub cmdPrevious_Click()
On Error Resume Next
 rs.MovePrevious
 If rs.BOF Then
  rs.MoveLast
 End If
 Call LoadData
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    If (txtDealerNo.Text = "") And (txtDealerName.Text = "") And (txtDealerContact.Text = "") And (txtDealerOutstanding.Text = "") Then
        MsgBox "Please Enter Sufficient Data", , "Cosmeta"
    Else
    rs!DealerID = txtDealerNo.Text
    rs!DealerName = txtDealerName.Text
    rs!DealerPhoneno = txtDealerContact.Text
    rs!DealerAddress = txtDealerAddress.Text
    rs!Email = txtEmail.Text
    rs!Outstandings = txtDealerOutstanding.Text
    rs!Details = txtDetails.Text
    rs.Update
    End If
End Sub



Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Module1.Connect
Module1.retData "Select * from Dealer"
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
txtDealerNo.Text = rs!DealerID
txtDealerName.Text = rs!DealerName
txtDealerContact.Text = rs!DealerPhoneno
txtDealerAddress.Text = rs!DealerAddress
txtEmail.Text = rs!Email
txtDealerOutstanding.Text = rs!Outstandings
txtDetails.Text = rs!Details
End Sub

Private Sub Image1_Click()
    On Error GoTo Trap
    Module1.retData "Select * from Dealer Where DealerName like '%" & txtDealerName.Text & "%' "
    Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
    Call LoadData
    cmdBack.Visible = True
Trap:
    Exit Sub
End Sub

Private Sub mnuAllDealer_Click()
DealerDataReport.Show
End Sub

Private Sub mnuDealerOutstanding_Click()
DataEnvironment1.DealerOutstanding
Load DealerOutstandingReport
DealerOutstandingReport.Show
DealerOutstandingReport.Refresh
DataEnvironment1.rsDealerOutstanding.Close
End Sub

Private Sub txtDealerAddress_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  txtEmail.SetFocus
 End If
End Sub

Private Sub txtDealerContact_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  txtDealerAddress.SetFocus
 End If
End Sub

Private Sub txtDealerContact_KeyUp(KeyCode As Integer, Shift As Integer)
If (Not IsNumeric(txtDealerContact.Text)) Then
 MsgBox "Enter Valid Price", vbExclamation, "Wrong Value Entered!!!"
 txtDealerContact.Text = ""
End If
End Sub

Private Sub txtDealerName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtDealerContact.SetFocus
End If

End Sub

Private Sub txtDealerNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtDealerName.SetFocus
End If
End Sub


Private Sub txtDealerOutstanding_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtDetails.SetFocus
End If
End Sub

Private Sub txtDealerOutstanding_KeyUp(KeyCode As Integer, Shift As Integer)
If (Not IsNumeric(txtDealerOutstanding.Text)) Then
 MsgBox "Enter Valid Price", vbExclamation, "Wrong Value Entered!!!"
 txtDealerOutstanding.Text = ""
End If
End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  txtDealerOutstanding.SetFocus
 End If
End Sub
