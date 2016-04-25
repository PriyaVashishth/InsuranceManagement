VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Application Form"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   14985
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14280
      Top             =   600
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   11880
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo2 
      Height          =   405
      ItemData        =   "CF.frx":0000
      Left            =   7080
      List            =   "CF.frx":0019
      TabIndex        =   35
      Text            =   " "
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "AGENTNAME"
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   1920
      TabIndex        =   34
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      DataField       =   "DATED"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   12360
      TabIndex        =   16
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      DataField       =   "AMTDEP"
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   12360
      TabIndex        =   15
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      DataField       =   "PROPNO"
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   12360
      TabIndex        =   14
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      DataField       =   "DOEXPIRY"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/M/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   7080
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "LICNSNO"
      DataSource      =   "Adodc2"
      Height          =   405
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   6135
      Left            =   0
      TabIndex        =   18
      Top             =   3360
      Width           =   15015
      Begin VB.CommandButton Command1 
         Caption         =   "View Plans"
         Height          =   375
         Left            =   12000
         TabIndex        =   53
         Top             =   4200
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   375
         Left            =   12000
         TabIndex        =   57
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         Height          =   375
         Left            =   12000
         TabIndex        =   58
         Top             =   4920
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         DataField       =   "GENDER"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   12720
         MaxLength       =   6
         TabIndex        =   56
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text15 
         DataField       =   "OBJECTINS"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   54
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox Text21 
         DataField       =   "AMTDEP"
         DataSource      =   "Adodc2"
         Height          =   405
         Left            =   8880
         TabIndex        =   52
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox Text20 
         DataField       =   "SUM"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   8880
         MaxLength       =   6
         TabIndex        =   50
         Top             =   4800
         Width           =   2295
      End
      Begin VB.TextBox Text19 
         DataField       =   "PLANTERM"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   8880
         MaxLength       =   2
         TabIndex        =   48
         Top             =   4200
         Width           =   2295
      End
      Begin VB.ComboBox Plans1 
         DataField       =   "PLAN"
         DataSource      =   "Adodc1"
         Height          =   405
         ItemData        =   "CF.frx":0053
         Left            =   7200
         List            =   "CF.frx":0066
         TabIndex        =   46
         Top             =   3720
         Width           =   6015
      End
      Begin VB.Frame Frame2 
         Caption         =   "Nominee"
         Height          =   2535
         Left            =   240
         TabIndex        =   36
         Top             =   3360
         Width           =   6015
         Begin VB.CommandButton Command3 
            Caption         =   "Load Picture"
            Height          =   375
            Left            =   3960
            TabIndex        =   60
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox Text18 
            DataField       =   "NOMADRS"
            DataSource      =   "Adodc1"
            Height          =   405
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   43
            Top             =   1440
            Width           =   2775
         End
         Begin VB.ComboBox Combo3 
            DataField       =   "NOMREL"
            DataSource      =   "Adodc1"
            Height          =   405
            ItemData        =   "CF.frx":009E
            Left            =   1920
            List            =   "CF.frx":00BA
            TabIndex        =   41
            Text            =   "Relation"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text8 
            DataField       =   "NOMAGE"
            DataSource      =   "Adodc1"
            Height          =   405
            Left            =   840
            MaxLength       =   2
            TabIndex        =   40
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox Text3 
            DataField       =   "NOMINAME"
            DataSource      =   "Adodc1"
            Height          =   405
            Left            =   840
            MaxLength       =   10
            TabIndex        =   38
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label26 
            Caption         =   "Photo"
            Height          =   255
            Left            =   4440
            TabIndex        =   44
            Top             =   240
            Width           =   855
         End
         Begin VB.Image Image1 
            Height          =   1215
            Left            =   4080
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Address"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label25 
            Caption         =   "Age"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox Text17 
         DataField       =   "NATION"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   9360
         MaxLength       =   5
         TabIndex        =   33
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text16 
         DataField       =   "BIRTHPLACE"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   6000
         MaxLength       =   7
         TabIndex        =   31
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text14 
         DataField       =   "PINRES"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   11640
         MaxLength       =   6
         TabIndex        =   28
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Text13 
         DataField       =   "PINAC"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   6960
         MaxLength       =   6
         TabIndex        =   25
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text12 
         DataField       =   "NAME"
         DataSource      =   "Adodc1"
         Height          =   855
         Left            =   240
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox Text11 
         DataField       =   "ADDRESSCOMM"
         DataSource      =   "Adodc1"
         Height          =   855
         Left            =   5160
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox Text10 
         DataField       =   "ADDRESSRES"
         DataSource      =   "Adodc1"
         Height          =   855
         Left            =   10320
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Exit"
         Height          =   375
         Left            =   12000
         TabIndex        =   59
         Top             =   5280
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Gender"
         Height          =   375
         Left            =   11400
         TabIndex        =   55
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Amount Deposited"
         Height          =   375
         Left            =   6840
         TabIndex        =   51
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label29 
         Caption         =   "Sum Proposed"
         Height          =   375
         Left            =   7200
         TabIndex        =   49
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label28 
         Caption         =   "Term of Plan"
         Height          =   375
         Left            =   7200
         TabIndex        =   47
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "Plans"
         Height          =   375
         Left            =   8160
         TabIndex        =   45
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Nationality"
         Height          =   375
         Left            =   8160
         TabIndex        =   32
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label23 
         Caption         =   "Place of Birth"
         Height          =   375
         Left            =   4440
         TabIndex        =   30
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label22 
         Caption         =   "Object of Insurance"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   15000
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   5040
         X2              =   5040
         Y1              =   120
         Y2              =   2400
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   10080
         X2              =   10080
         Y1              =   120
         Y2              =   2400
      End
      Begin VB.Label Label21 
         Caption         =   "Pin Code"
         Height          =   375
         Left            =   10440
         TabIndex        =   27
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Pin Code"
         Height          =   375
         Left            =   5760
         TabIndex        =   26
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Residential Address (if different)"
         Height          =   375
         Left            =   10320
         TabIndex        =   21
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Address to which communication is to be sent"
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label17 
         Caption         =   "Full name"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   12240
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from projins"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10680
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from customers"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Line Line10 
      X1              =   0
      X2              =   0
      Y1              =   1320
      Y2              =   3000
   End
   Begin VB.Line Line4 
      X1              =   15000
      X2              =   15000
      Y1              =   1320
      Y2              =   3000
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   240
      Picture         =   "CF.frx":00FD
      Top             =   120
      Width           =   1980
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "All answers to be filled in Legibly, Answers must be given in words"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   3120
      Width           =   6375
   End
   Begin VB.Label Label15 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   13
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Amount of Deposit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Proposal No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "For Office Use"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Date of Expiry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Branch Office"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Line Line6 
      X1              =   9360
      X2              =   9360
      Y1              =   1320
      Y2              =   3000
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   15000
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line3 
      X1              =   4080
      X2              =   15000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label7 
      Caption         =   "Licence No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Agent's Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shimla"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Division"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   4440
      X2              =   4440
      Y1              =   1320
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4080
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "(Not to be used for insurance on the lives of minors)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Proposal For Insurance on own Life "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sheffi As Boolean, priya As Boolean
Private Sub Command1_Click()
Command1.ToolTipText = "View the plans available"
Form3.Show
 Form2.Hide
End Sub

Private Sub Command2_Click()
Adodc1.Enabled = True
Adodc2.Enabled = True
Adodc1.Recordset.Update
Adodc2.Recordset.Update
End Sub


Private Sub Command3_Click()
cd.ShowOpen
Image1.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1_Timer
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Timer1.Interval = 0
Adodc1.Enabled = False
Adodc2.Enabled = False
    Text9.Text = Now
    Adodc1.Refresh
    Adodc2.Refresh
    Adodc1.Recordset.AddNew
    Adodc2.Recordset.AddNew
sheffi = True
priya = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Interval = 0
MDIForm1.StatusBar1.Panels(1).Text = ""
sheffi = True
priya = True
End Sub

Private Sub Text15_GotFocus()
If Text10.Text = "" And Text14.Text = "" Then
    Text10.Text = Text11.Text
    Text14.Text = Text13.Text
End If
End Sub

Private Sub Timer1_Timer()
If sheffi Then
    MDIForm1.StatusBar1.Panels(1).Text = "Click this button to upload the nominee's picture"
    sheffi = False
    priya = True
ElseIf priya Then
    priya = False
    sheffi = True
    MDIForm1.StatusBar1.Panels(1).Text = ""
End If
End Sub
