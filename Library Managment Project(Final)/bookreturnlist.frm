VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form bookreturnlist 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "bookreturnlist.frx":0000
   ScaleHeight     =   8550
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "CARD DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Select an option from the list to view details"
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "CARDNO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7320
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   7080
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "bookreturnlist.frx":3A95
      Height          =   5295
      Left            =   1320
      OleObjectBlob   =   "bookreturnlist.frx":3AA9
      TabIndex        =   0
      Top             =   1680
      Width           =   9555
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ISSUEBOOK"
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH RESULT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   0
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      Width           =   4005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL NO OF ENTRIES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   7080
      Width           =   2730
   End
End
Attribute VB_Name = "bookreturnlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
databasedetails.Show
Unload Me
End Sub

Private Sub Command2_Click()
singlecarddetail.Data1.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text2.Text & "' or CARDNO2='" & Text2.Text & "'"
singlecarddetail.Data1.Refresh
singlecarddetail.Show
End Sub
