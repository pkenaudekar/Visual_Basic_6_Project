VERSION 5.00
Begin VB.Form newaccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creat an account"
   ClientHeight    =   2460
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin2.frx":0000
   ScaleHeight     =   1453.449
   ScaleMode       =   0  'User
   ScaleWidth      =   4112.583
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      DataField       =   "ADMINACC"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3600
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C00000&
      Caption         =   "Administrator Account"
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
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      ToolTipText     =   "Check this option to provide administrative rights"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "PASSWORD"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ACCOUNT"
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ACCOUNT"
      Top             =   360
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.TextBox Text1 
      DataField       =   "PASSWORD"
      DataSource      =   "Data1"
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
      Left            =   1920
      PasswordChar    =   "|"
      TabIndex        =   6
      ToolTipText     =   "Re-enter password"
      Top             =   1080
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      DataField       =   "USERNAME"
      DataSource      =   "Data1"
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
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Enter username"
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "CREATE"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Creates a new account"
      Top             =   1920
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Goes back to the previous page"
      Top             =   1920
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
      Left            =   1920
      PasswordChar    =   "|"
      TabIndex        =   3
      ToolTipText     =   "Enter password"
      Top             =   600
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Enter Password:"
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
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
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
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password:"
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
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   600
      Width           =   1410
   End
End
Attribute VB_Name = "newaccount"
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
MsgBox "Enter Your Username", , "Error"
ElseIf txtPassword.Text = "" Then
MsgBox "Enter Your Password", , "Error"
ElseIf Text1.Text = "" Then
MsgBox "Re-Enter Your Password", , "Error"
Else
    If Text4.Text = "Y" Then
        Text2.Text = txtUserName.Text
        If txtPassword.Text = Text1.Text Then
        Dim strfind As String
        Dim intlen As Integer
        strfind = txtPassword.Text
        intlen = Len(strfind)
            If intlen < 6 Then
            MsgBox "The Password Should Be Atleast Six Character Long ", , "Error"
            txtPassword.Text = ""
            Text1.Text = ""
            Else
            newaccount.Data2.RecordSource = "SELECT * FROM ACCOUNT WHERE USERNAME = '" & Text2.Text & "' "
            newaccount.Data2.Refresh
                If Text3.Text = "" Then
                    If Check1.Value = 1 Then
                    Text5.Text = "Y"
                    Else
                    Text5.Text = "N"
                    End If
                Data1.Recordset.AddNew
                MsgBox "New Account Was Successfull Created", , ""
                txtUserName.Text = ""
                txtPassword.Text = ""
                Text1.Text = ""
                Check1.Value = 0
                Else
                MsgBox "The Username Already Exsists.Please Select A New Username", , "Error"
                txtUserName.Text = ""
                txtPassword.Text = ""
                Text1.Text = ""
                Text2.Text = ""
                Text3.Text = ""
                Check1.Value = 0
                End If
            End If
        Else
        MsgBox "The Passwords You Typed Do Not Match.Please Re-Type The New Password In Both Boxes ", , "Error"
        txtPassword.Text = ""
        Text1.Text = ""
        End If
    End If
End If

End Sub

