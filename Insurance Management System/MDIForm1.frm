VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000011&
   Caption         =   "Way to go.."
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13995
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      Height          =   7935
      Left            =   0
      Picture         =   "MDIForm1.frx":0000
      ScaleHeight     =   7875
      ScaleWidth      =   13935
      TabIndex        =   1
      Top             =   0
      Width           =   13995
      Begin VB.PictureBox Picture2 
         Height          =   1095
         Index           =   0
         Left            =   0
         Picture         =   "MDIForm1.frx":6B0D
         ScaleHeight     =   1035
         ScaleWidth      =   1875
         TabIndex        =   2
         Top             =   4920
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7815
      Visible         =   0   'False
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuIns 
      Caption         =   "Insurance"
      Begin VB.Menu mnuNM 
         Caption         =   "New Proposal"
      End
      Begin VB.Menu mnuv 
         Caption         =   "View Proposal Details"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuNM_Click()
   Form2.Show
   Picture1.Visible = False
End Sub

Private Sub mnuv_Click()
   Picture1.Visible = False
   Form1.Show
End Sub

