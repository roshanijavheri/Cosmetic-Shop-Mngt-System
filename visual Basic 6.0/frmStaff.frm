VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Details"
   ClientHeight    =   9105
   ClientLeft      =   4620
   ClientTop       =   1560
   ClientWidth     =   15885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   15885
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
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbGender 
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
      ItemData        =   "frmStaff.frx":0000
      Left            =   2880
      List            =   "frmStaff.frx":000D
      TabIndex        =   26
      Top             =   6240
      Width           =   3495
   End
   Begin VB.TextBox txtEmpNo 
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
      Left            =   2880
      TabIndex        =   14
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtEmpSalary 
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
      Left            =   2880
      TabIndex        =   13
      Top             =   5280
      Width           =   3495
   End
   Begin VB.TextBox txtEmpDesignation 
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
      Left            =   2880
      TabIndex        =   12
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox txtEmpAdd 
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
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox txtEmpContact 
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
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtEmpName 
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
      Left            =   2880
      TabIndex        =   9
      Top             =   1440
      Width           =   3375
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
      Left            =   6840
      TabIndex        =   8
      Top             =   8280
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
      Left            =   11160
      TabIndex        =   7
      Top             =   8280
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
      Left            =   13440
      TabIndex        =   6
      Top             =   8280
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
      Left            =   9120
      TabIndex        =   5
      Top             =   8280
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
      Left            =   11400
      TabIndex        =   4
      Top             =   7320
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
      Left            =   10080
      TabIndex        =   3
      Top             =   7320
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
      Left            =   7920
      TabIndex        =   2
      Top             =   7320
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
      Left            =   12720
      TabIndex        =   1
      Top             =   7320
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6975
      Left            =   6600
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12303
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
   Begin MSComCtl2.DTPicker dtpDOB 
      Height          =   615
      Left            =   2880
      TabIndex        =   15
      Top             =   7080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      _Version        =   393216
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   138477571
      CurrentDate     =   44247
   End
   Begin MSComCtl2.DTPicker dtpDOJ 
      Height          =   615
      Left            =   2880
      TabIndex        =   16
      Top             =   8040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      _Version        =   393216
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe MDL2 Assets"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   138477571
      CurrentDate     =   44247
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   5520
      Picture         =   "frmStaff.frx":0026
      Stretch         =   -1  'True
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Joining"
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
      TabIndex        =   25
      Top             =   8160
      Width           =   2445
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
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
      TabIndex        =   24
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label8 
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
      Left            =   600
      TabIndex        =   23
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
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
      TabIndex        =   22
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emp. Name"
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
      TabIndex        =   21
      Top             =   1440
      Width           =   1830
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Base Salary"
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
      TabIndex        =   20
      Top             =   5280
      Width           =   1770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   720
      TabIndex        =   19
      Top             =   6240
      Width           =   1155
   End
   Begin VB.Label Label3 
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
      Left            =   600
      TabIndex        =   18
      Top             =   2400
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
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
      TabIndex        =   17
      Top             =   4320
      Width           =   1920
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAllStaff 
         Caption         =   "All Staff Report"
      End
      Begin VB.Menu mnuDesignation 
         Caption         =   "Reports According Designation"
      End
      Begin VB.Menu mnuSalary 
         Caption         =   "According Salary"
         Begin VB.Menu mnuGreterSalary 
            Caption         =   "Salary Greater Than"
         End
         Begin VB.Menu mnuLessSalary 
            Caption         =   "Salary Less Than"
         End
      End
      Begin VB.Menu mnuID 
         Caption         =   "Report Having Id"
      End
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New Recordset

Private Sub cmdAddNew_Click()
    Call cmdClear_Click
    rs.AddNew
    rs1.Open "Select Max(EmployeeID) from Staff", Con, adOpenStatic, adLockOptimistic
    txtEmpNo.Text = rs1.Fields(0).Value + 1
    rs1.Close
End Sub

Private Sub cmdBack_Click()
Module1.retData "Select * from Staff"
Call LoadData
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
cmdBack.Visible = False
End Sub

Private Sub cmdClear_Click()
    txtEmpNo.Text = ""
    txtEmpName.Text = ""
    txtEmpContact.Text = ""
    txtEmpAdd.Text = ""
    cmbGender.Text = "Select Gender"
    txtEmpDesignation.Text = ""
    txtEmpSalary.Text = ""
    dtpDOB.Refresh
    dtpDOJ.Refresh
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim Confirm As String
Confirm = MsgBox("Are You Sure, You Want To delete this " & rs!EmployeeName & " Record ?", vbYesNo + vbCritical, "Delete Record Confirmation!!!")
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
  If (txtEmpNo.Text = "") And (txtEmpName.Text = "") And (txtEmpContact.Text = "") And (txtEmpAdd.Text = "") And (txtEmpDesignation.Text = "") And (txtEmpSalary.Text = "") Then
    MsgBox "Please Enter Sufficient Data", , "Cosmeta"
  Else
   rs.Fields(0) = txtEmpNo.Text
   rs.Fields("EmployeeName") = txtEmpName.Text
   rs.Fields("EmployeeAddress") = txtEmpAdd.Text
   rs.Fields("EmployeePhoneno") = txtEmpContact.Text
   rs.Fields("Designation") = txtEmpDesignation.Text
   rs.Fields("BasicSalary") = txtEmpSalary.Text
   rs.Fields("Gender") = cmbGender.Text
   rs.Fields("DateOfBirth") = dtpDOB.Value
   rs.Fields("JoiningDate") = dtpDOJ.Value
   rs.Update
 End If
End Sub


Private Sub txtDOB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDOJ.SetFocus
    End If
End Sub


Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
 Module1.Connect
 Module1.retData "Select * from Staff"
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
   txtEmpNo.Text = rs.Fields(0)
   txtEmpName.Text = rs.Fields("EmployeeName")
   txtEmpAdd.Text = rs.Fields("EmployeeAddress")
   txtEmpContact.Text = rs.Fields("EmployeePhoneno")
   txtEmpDesignation.Text = rs.Fields("Designation")
   txtEmpSalary.Text = rs.Fields("BasicSalary")
   cmbGender.Text = rs.Fields("Gender")
   dtpDOB.Value = rs.Fields("DateOfBirth")
   dtpDOJ.Value = rs.Fields("JoiningDate")
   
End Sub

Private Sub Image1_Click()
    On Error GoTo Trap
    Module1.retData "Select * from Staff Where EmployeeName like '%" & txtEmpName.Text & "%' "
    Set DataGrid1.DataSource = rs
    DataGrid1.Refresh
    Call LoadData
    cmdBack.Visible = True
Trap:
    Exit Sub
End Sub

Private Sub mnuAllStaff_Click()
StaffDataReport.Show
End Sub

Private Sub mnuDesignation_Click()       'Dynamic Staff Designation Report
If (txtEmpDesignation.Text = "") Then
 MsgBox "Please Enter Designation", vbInformation, "Cosmeta"
Else
    DataEnvironment1.StaffOnDesignation txtEmpDesignation.Text
    Load StaffonDesignationReport
    StaffonDesignationReport.Show
    StaffonDesignationReport.Refresh
    DataEnvironment1.rsStaffOnDesignation.Close
End If
End Sub

Private Sub mnuGreterSalary_Click()
If (txtEmpSalary.Text = "") Then
 MsgBox "Please Enter Salary to Search", vbInformation, "Cosmeta"
Else
    DataEnvironment1.StaffOnSalaryGraterThan txtEmpSalary.Text
    Load StaffSalaryGreaterThanDataReport
    StaffSalaryGreaterThanDataReport.Show
    StaffSalaryGreaterThanDataReport.Refresh
    DataEnvironment1.rsStaffOnSalaryGraterThan.Close
End If
End Sub

Private Sub mnuID_Click()
If (txtEmpNo.Text = "") Then
 MsgBox "Please Enter Employee ID", vbInformation, "Cosmeta"
Else
    DataEnvironment1.StaffHavingID txtEmpNo.Text
    Load StaffHavingIDDataReport
    StaffHavingIDDataReport.Show
    StaffHavingIDDataReport.Refresh
    DataEnvironment1.rsStaffHavingID.Close
End If
End Sub

Private Sub mnuLessSalary_Click()
If (txtEmpSalary.Text = "") Then
 MsgBox "Please Enter Salary", vbInformation, "Cosmeta"
Else
    DataEnvironment1.StaffonSalaryLessThan txtEmpSalary.Text
    Load StaffSalaryLessThanDataReport
    StaffSalaryLessThanDataReport.Show
    StaffSalaryLessThanDataReport.Refresh
    DataEnvironment1.rsStaffonSalaryLessThan.Close
End If
End Sub

Private Sub txtEmpContact_KeyUp(KeyCode As Integer, Shift As Integer)
If (Not IsNumeric(txtEmpContact.Text)) Then
 MsgBox "Enter Valid Price", vbExclamation, "Wrong Value Entered!!!"
 txtEmpContact.Text = ""
End If
End Sub

Private Sub txtEmpDesignation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEmpSalary.SetFocus
    End If
End Sub
Private Sub txtEmpAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtEmpDesignation.SetFocus
End If
End Sub

Private Sub txtEmpContact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtEmpAdd.SetFocus
End If
End Sub

Private Sub txtEmpName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtEmpContact.SetFocus
End If

End Sub

Private Sub txtEmpNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtEmpName.SetFocus
End If

End Sub


Private Sub txtEmpSalary_KeyUp(KeyCode As Integer, Shift As Integer)
If (Not IsNumeric(txtEmpSalary.Text)) Then
 MsgBox "Enter Valid Price", vbExclamation, "Wrong Value Entered!!!"
 txtEmpSalary.Text = ""
End If
End Sub
