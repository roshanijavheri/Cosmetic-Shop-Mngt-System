VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOrder 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9195
   ClientLeft      =   4620
   ClientTop       =   1470
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   15180
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   15690
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   882
      BackColor       =   -2147483626
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Sales Order"
      TabPicture(0)   =   "frmOrder.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdRemove"
      Tab(0).Control(1)=   "cmdSave"
      Tab(0).Control(2)=   "txtTotalBill"
      Tab(0).Control(3)=   "Text3"
      Tab(0).Control(4)=   "List1(4)"
      Tab(0).Control(5)=   "txtQty"
      Tab(0).Control(6)=   "List1(3)"
      Tab(0).Control(7)=   "List1(2)"
      Tab(0).Control(8)=   "Combo1"
      Tab(0).Control(9)=   "List1(0)"
      Tab(0).Control(10)=   "Text1"
      Tab(0).Control(11)=   "Text2"
      Tab(0).Control(12)=   "DTPicker1"
      Tab(0).Control(13)=   "List1(1)"
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(16)=   "Label6"
      Tab(0).Control(17)=   "Label5"
      Tab(0).Control(18)=   "Label4"
      Tab(0).Control(19)=   "Label1"
      Tab(0).Control(20)=   "Label2"
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Purchase Order"
      TabPicture(1)   =   "frmOrder.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label12"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblTotalProducts"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblTotalItem"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblTotalItemsValue"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "POdate"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Combo2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Combo3"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "List2(2)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "List2(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "List2(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtPQty"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmdPOSave"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtPOID"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmdPoRemove"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.CommandButton cmdPoRemove 
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
         Height          =   735
         Left            =   11520
         TabIndex        =   39
         Top             =   4560
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
         Height          =   735
         Left            =   -62160
         TabIndex        =   38
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtPOID 
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   33
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdPOSave 
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
         Left            =   11520
         TabIndex        =   32
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox txtPQty 
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
         Left            =   6600
         TabIndex        =   31
         ToolTipText     =   "Enter Quantity"
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         DataField       =   "ProductName"
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
         ItemData        =   "frmOrder.frx":0038
         Left            =   1320
         List            =   "frmOrder.frx":003A
         TabIndex        =   27
         Top             =   3600
         Width           =   7695
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmOrder.frx":003C
         Left            =   360
         List            =   "frmOrder.frx":003E
         TabIndex        =   26
         Top             =   3600
         Width           =   975
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmOrder.frx":0040
         Left            =   9000
         List            =   "frmOrder.frx":0042
         TabIndex        =   25
         Top             =   3600
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
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
         ItemData        =   "frmOrder.frx":0044
         Left            =   2640
         List            =   "frmOrder.frx":0046
         TabIndex        =   24
         Text            =   "Select Dealer"
         ToolTipText     =   "Click Here To Select Dealer"
         Top             =   1560
         Width           =   7575
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frmOrder.frx":0048
         Left            =   480
         List            =   "frmOrder.frx":004A
         TabIndex        =   22
         Text            =   "Select Products"
         Top             =   2400
         Width           =   5415
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
         Left            =   -64440
         TabIndex        =   20
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtTotalBill 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -63960
         TabIndex        =   19
         Top             =   8040
         Width           =   1695
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
         Height          =   540
         Left            =   -67200
         TabIndex        =   18
         Tag             =   "For Your Help"
         ToolTipText     =   "Enter Quantity"
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Left            =   -63960
         TabIndex        =   17
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtQty 
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
         Left            =   -69600
         TabIndex        =   11
         ToolTipText     =   "Enter Quantity"
         Top             =   2280
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmOrder.frx":004C
         Left            =   -66000
         List            =   "frmOrder.frx":004E
         TabIndex        =   10
         Top             =   3720
         Width           =   2055
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
         Left            =   -67800
         TabIndex        =   6
         Top             =   3720
         Width           =   1815
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
         ItemData        =   "frmOrder.frx":0050
         Left            =   -74640
         List            =   "frmOrder.frx":0052
         TabIndex        =   4
         Text            =   "Select Products"
         Top             =   2280
         Width           =   4935
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Left            =   -74640
         TabIndex        =   3
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -72840
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -72840
         TabIndex        =   1
         Top             =   1320
         Width           =   7455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   615
         Left            =   -65280
         TabIndex        =   7
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
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
         Format          =   138412035
         CurrentDate     =   44247
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Left            =   -73680
         TabIndex        =   5
         Top             =   3720
         Width           =   5895
      End
      Begin MSComCtl2.DTPicker POdate 
         Height          =   735
         Left            =   11040
         TabIndex        =   23
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
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
         Format          =   138412035
         CurrentDate     =   44247
      End
      Begin VB.Label lblTotalItemsValue 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   9480
         TabIndex        =   37
         Top             =   7800
         Width           =   1050
      End
      Begin VB.Label lblTotalItem 
         AutoSize        =   -1  'True
         Caption         =   "Total Item ="
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7200
         TabIndex        =   36
         Top             =   7800
         Width           =   2385
      End
      Begin VB.Label lblTotalProducts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   570
         Left            =   360
         TabIndex        =   35
         Top             =   7800
         Width           =   10695
      End
      Begin VB.Label Label12 
         Caption         =   "P_Order No :"
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   34
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label11 
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
         Left            =   360
         TabIndex        =   30
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label10 
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
         Left            =   1320
         TabIndex        =   29
         Top             =   3120
         Width           =   7695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quantity"
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
         Height          =   495
         Left            =   9000
         TabIndex        =   28
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Dealer Name :"
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   2295
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
         Left            =   -65640
         TabIndex        =   16
         Top             =   8160
         Width           =   1695
      End
      Begin VB.Label Label7 
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
         Height          =   735
         Left            =   -67560
         TabIndex        =   15
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -65640
         TabIndex        =   14
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label5 
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
         Height          =   735
         Left            =   -71640
         TabIndex        =   13
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         Height          =   735
         Left            =   -74280
         TabIndex        =   12
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Order No :"
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74640
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Customer Name :"
         BeginProperty Font 
            Name            =   "Microsoft YaHei"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74640
         TabIndex        =   8
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Menu mnuOrders 
      Caption         =   "Orders"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuPurchaseOrder 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu mnuSalesOrder 
         Caption         =   "Sales Orders"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuProductRepo 
         Caption         =   "Product Reports"
      End
      Begin VB.Menu mnuBill 
         Caption         =   "Billing Report"
      End
      Begin VB.Menu mnuDealerReport 
         Caption         =   "Dealer Reports"
      End
      Begin VB.Menu mnuCustomerReports 
         Caption         =   "Customer Report"
      End
      Begin VB.Menu mnuStaffReports 
         Caption         =   "Staff Reports"
      End
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Change the Code Because Adodc has been Deleted
Dim i, j, CurrentAmount As Integer
Dim rs1 As New Recordset

Private Sub cmdPoRemove_Click()
If (Not (List2(2).ListIndex < 0)) Then
List2(2).RemoveItem List2(1).ListIndex
List2(1).RemoveItem List2(1).ListIndex
List2(0).Clear
For itr = 0 To List2(1).ListCount - 1
  List2(0).AddItem itr + 1
Next
j = j - 1
 For itr = 0 To List2(2).ListCount - 1
       TotalQty = TotalQty + Val(List2(2).List(itr))
      Next
        lblTotalItemsValue.Caption = TotalQty
Else
 MsgBox "Nothing To remove"
End If
End Sub

Private Sub cmdPOSave_Click()                                          'Save Data into Database for Purchase Order
Dim i As Integer
    Module1.retData "Select * from PurchaseOrder"
    rs.AddNew
    For i = 0 To List2(1).ListCount - 1
        rs.Fields("POID") = txtPOID.Text
        rs.Fields("PODate") = POdate.Value
        rs.Fields("DealerName") = Combo3
        rs.Fields("ProductName") = List2(1).List(i)
        rs.Fields("POQuantity") = List2(2).List(i)
        rs.Update
        rs.AddNew
    Next
    rs.Fields("POID") = txtPOID.Text
    rs.Fields("PODate") = POdate.Value
    rs.Fields("DealerName") = Combo3
    rs.Fields("POQuantity") = lblTotalItemsValue.Caption
    rs.Update
    Call Clear
    rs.AddNew
    
    MsgBox "Purchase Order Saved Successfully" & vbNewLine & "Go to Purchase Order For Details"

End Sub

Private Sub cmdRemove_Click()
List1(2).RemoveItem List1(1).ListIndex
CurrentAmount = CurrentAmount - (List1(4).List(List1(1).ListIndex))   ' Subtract Amount of Product Removed
txtTotalBill.Text = CurrentAmount                                     ' For Displaying The TotalAmount
List1(3).RemoveItem List1(1).ListIndex
List1(4).RemoveItem List1(1).ListIndex
List1(1).RemoveItem List1(1).ListIndex
List1(0).Clear
For itr = 0 To List1(1).ListCount - 1
  List1(0).AddItem itr + 1
Next
i = i - 1
End Sub

Private Sub cmdSave_Click()                                           'Save Data into Database for Sales Order
    Dim i As Integer
    Module1.retData "Select * from SalesOrder"
    rs.AddNew
    For i = 0 To List1(1).ListCount - 1
        rs.Fields("OrderId") = Text1.Text
        rs.Fields("CustomerName") = Text2.Text
        rs.Fields("ProductName") = List1(1).List(i)
        rs.Fields("SellingPrice") = List1(2).List(i)
        rs.Fields("Qty") = List1(3).List(i)
        rs.Fields("TotalAmount") = "0"
        rs.Update
        rs.AddNew
    Next
    rs.Fields("CustomerName") = Text2.Text
    rs.Fields("TotalAmount") = txtTotalBill.Text
    rs.Update
    Call Clear
    rs.AddNew
End Sub

Private Sub Clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
txtQty.Text = ""
Combo1.Text = "Select Product"
List1.Item(0).Clear
List1.Item(1).Clear
List1.Item(2).Clear
List1.Item(3).Clear
List1.Item(4).Clear
txtTotalBill.Text = ""
End Sub


Private Sub Combo1_Click()
On Error Resume Next
Dim Amount1, NewQty As Integer
If Combo1.Text <> "" Then
Module1.retData "Select * from Product Where ProductName='" & Combo1 & "' "
'Don't Remove This Code. This is For Duplicate Product Entries{
 For itr = 0 To List1(1).ListCount - 1
    If (Combo1.Text = List1(1).List(itr)) Then
      NewQty = InputBox("You Have Selected " & Combo1.Text & "again" & vbNewLine & "Enter New Qty", "Same Product Selected!!!")
      If (NewQty <= 0) Then
       Exit Sub
      Else
      List1(3).List(itr) = NewQty
      List1(4).List(itr) = NewQty * List1(2).List(itr)
       For itr1 = 0 To List1(4).ListCount - 1
        Amount1 = Amount1 + Val(List1(4).List(itr1))
        CurrentAmount = Amount1
        txtTotalBill.Text = CurrentAmount
       Next
      Combo1.SetFocus
      Exit Sub
      End If
      End If
 Next itr
 
'Dont Remove - VIMP Code}
 
List1(0).AddItem i                                   'Code For Sr No
i = i + 1
List1(1).AddItem Combo1.Text                               ' Code For Product Name In List Box
List1(2).AddItem rs.Fields("SellingPrice")             ' Code For Amount Section
Text3.Text = rs.Fields("SellingPrice")                    'this is useful for Converting Selling Price in Integer Format
AvailableStock = rs.Fields("StockLevel")
txtQty.SetFocus
End If
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
Module1.retData "Select ProductName from Product where ProductName Like '" & Combo1.Text & "' "
End Sub

Private Sub Combo2_Click()
On Error GoTo trap
Dim NewQty1 As Integer
If Combo2.Text <> "" Then
Module1.retData "Select * from Product Where ProductName='" & Combo2 & "' "
'Don't Remove This Code. This is For Duplicate Product Entries{
 For itr = 0 To List2(1).ListCount - 1
    If (Combo2.Text = List2(1).List(itr)) Then
      NewQty1 = InputBox("You Have Selected " & Combo2.Text & "again" & vbNewLine & "Enter New Qty", "Same Product Selected!!!")
      If (NewQty1 <= 0) And (Not (IsNumeric(NewQty))) Then
       MsgBox "Enter Valid Quantity", vbExclamation, "Change Quantity"
       Exit Sub
      Else
      List2(2).List(itr) = NewQty1
      For itr1 = 0 To List2(2).ListCount - 1
       TotalQty = TotalQty + Val(List2(2).List(itr1))
      Next
        lblTotalItemsValue.Caption = TotalQty
      Combo2.SetFocus
      Exit Sub
      End If
      End If
 Next itr
 
'Dont Remove - VIMP Code}
 
List2(0).AddItem j                                   'Code For Sr No
j = j + 1
List2(1).AddItem Combo2.Text                               ' Code For Product Name In List Box
txtPQty.SetFocus
End If


trap:
Exit Sub
End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
Module1.retData "Select ProductName from Product where ProductName Like '" & Combo2.Text & "'  "
End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
     j = 1
     i = 1 ' For Product Sr Column
  Module1.Connect
  Module1.retData "Select * from Product"
   With rs           ' Code for Loading Database into Combo box
    Do Until .EOF
     Combo1.AddItem ![ProductName]
     Combo2.AddItem ![ProductName]
     .MoveNext
    Loop
  End With
   rs1.Open ("Select Max(OrderId) from SalesOrder"), Con, adOpenStatic, adLockOptimistic
 Text1.Text = rs1.Fields(0).Value + 1
 DTPicker1.Value = Date
 POdate.Value = Date
 
End Sub



Private Sub List1_Click(Index As Integer)      ' For Selecting Whole Row at a time
List1(0).ListIndex = List1(1).ListIndex
List1(2).ListIndex = List1(1).ListIndex
List1(3).ListIndex = List1(1).ListIndex
List1(4).ListIndex = List1(1).ListIndex
End Sub

Private Sub List2_Click(Index As Integer)
    ' For Selecting Whole Row at once
List2(0).ListIndex = List2(1).ListIndex
List2(1).ListIndex = List2(1).ListIndex
List2(2).ListIndex = List2(1).ListIndex

End Sub

Private Sub mnuBill_Click()
BillConvinienceReport.Show
End Sub

Private Sub mnuCustomerReports_Click()
CustomerDataReport.Show
End Sub

Private Sub mnuDealerReport_Click()
DealerDataReport.Show
End Sub

Private Sub mnuProductRepo_Click()
ProductReport.Show
End Sub


Private Sub mnuPurchaseOrder_Click()
frmAllPO.Show
End Sub

Private Sub mnuSalesOrder_Click()
frmAllSalesOrder.Show
End Sub

Private Sub mnuStaffReports_Click()
StaffDataReport.Show
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Module1.retData "Select * from Dealer"
 With rs           ' Code for Loading Dealer Database into Combo box
    Do Until .EOF
     Combo3.AddItem ![DealerName]
     .MoveNext
    Loop
  End With
  
Module1.retData "Select Max(POID) from PurchaseOrder"
txtPOID.Text = rs.Fields(0) + 1

End Sub


Private Sub txtPQty_KeyPress(KeyAscii As Integer)
If List2(1).ListCount > List2(2).ListCount Then
Dim TotalQty As Integer    ' For Total Qty in PO
If KeyAscii = 13 Then
     If (Val(txtPQty.Text) <= 0) Then
      MsgBox "You Have Entered" & txtPQty.Text, vbCritical + vbExclamation, "Cosmeta"
      Exit Sub
      End If
      List2(2).AddItem (txtPQty.Text)
      For itr = 0 To List2(2).ListCount - 1
       TotalQty = TotalQty + Val(List2(2).List(itr))
       lblTotalItemsValue.Caption = TotalQty
      Next
      lblTotalProducts.Caption = "    Total Products = " & List2(2).ListCount
      txtPQty.Text = ""
      Combo2.SetFocus
End If
End If
End Sub


Private Sub txtQty_KeyPress(KeyAscii As Integer)
Dim Calculate As Integer
 If KeyAscii = 13 Then
    List1(3).AddItem (txtQty.Text)
    List1(4).AddItem (Val(Text3.Text) * Val(txtQty.Text))
    Calculate = (Val(Text3.Text) * Val(txtQty.Text))        ' For Converting The List Item Into Integer
    CurrentAmount = CurrentAmount + Calculate         ' Change this if the Item is Deleted
    txtTotalBill.Text = CurrentAmount
    txtQty.Text = ""
    Combo1.SetFocus
 End If
End Sub
