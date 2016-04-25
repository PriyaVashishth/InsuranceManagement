VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Plans"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   15360
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5040
      Top             =   7320
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
      RecordSource    =   "select * from plans"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back To Form"
      Height          =   615
      Left            =   7080
      TabIndex        =   16
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm this Plan"
      Height          =   615
      Left            =   9840
      TabIndex        =   8
      Top             =   7080
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   7680
      Picture         =   "WLP.frx":0000
      ScaleHeight     =   6615
      ScaleWidth      =   9615
      TabIndex        =   17
      Top             =   840
      Width           =   9615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "CRITILLRIDER"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4560
      TabIndex        =   15
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Critical Illness Rider"
      Height          =   495
      Left            =   720
      TabIndex        =   14
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "NETYIELD"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4560
      TabIndex        =   13
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "MATURITYAGE"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ACCIDENTBENEFIT"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4560
      TabIndex        =   11
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "MMTERM"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "MMAGE"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4560
      TabIndex        =   9
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Plan's Name"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Yield(%)"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Accident Benefit(%)"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Max. maturity age"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Min./Max. Term"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Min./Max. age at entry"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Plans 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Plans"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a As String

Private Sub Command1_Click()
    a = Label6.Caption
End Sub

Private Sub Command2_Click()
    Form2.Show
    Unload Form3
    Form2.Plans1.Text = a
End Sub

