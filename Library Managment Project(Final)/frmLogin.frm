VERSION 5.00
Begin VB.Form stafflogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1530
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   903.974
   ScaleMode       =   0  'User
   ScaleWidth      =   3605.552
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      DataField       =   "ADMINACC"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3840
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataField       =   "USERNAME"
      DataSource      =   "Data1"
      Height          =   345
      Left            =   5520
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "PASSWORD"
      DataSource      =   "Data1"
      Height          =   345
      Left            =   3840
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "USERNAME"
      DataSource      =   "Data1"
      Height          =   345
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ACCOUNT"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Enter your username"
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "LOGIN"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "|"
      TabIndex        =   3
      ToolTipText     =   "Enter your password"
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "stafflogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()

If txtUserName.Text = "" Then
MsgBox "Please enter a username", , "Error"
ElseIf txtPassword.Text = "" Then
MsgBox "Please enter a password", , "Error"
Else
Data1.RecordSource = "SELECT * FROM ACCOUNT WHERE USERNAME = '" & txtUserName.Text & "' "
Data1.Refresh
    If Text1.Text = "" Or Text2.Text = "" Then
    MsgBox "Invalid Username or Password, try again!", , "Error"
    SendKeys "{Home}+{End}"
    stafflogin.txtUserName.Text = ""
    stafflogin.txtPassword.Text = ""
    stafflogin.Text1.Text = ""
    stafflogin.Text2.Text = ""
    stafflogin.Text4.Text = ""
    ElseIf txtUserName.Text = Text1.Text And txtPassword.Text = Text2.Text Then
    optionpage.Show
    Unload startuppage
    LoginSucceeded = True
    Me.Hide
    Else
    MsgBox "Invalid Username or Password, try again!", , "Error"
    SendKeys "{Home}+{End}"
    stafflogin.txtUserName.Text = ""
    stafflogin.txtPassword.Text = ""
    stafflogin.Text1.Text = ""
    stafflogin.Text2.Text = ""
    stafflogin.Text4.Text = ""
    End If
End If
     
End Sub

