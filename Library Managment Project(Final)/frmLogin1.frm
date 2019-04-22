VERSION 5.00
Begin VB.Form changepassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2100
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin1.frx":0000
   ScaleHeight     =   1240.75
   ScaleMode       =   0  'User
   ScaleWidth      =   4436.519
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ACCOUNT"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
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
      Left            =   4680
      TabIndex        =   9
      Top             =   550
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
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
      Left            =   4680
      TabIndex        =   8
      Top             =   144
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text3 
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
      Left            =   2400
      PasswordChar    =   "|"
      TabIndex        =   7
      ToolTipText     =   "Re-Enter your new password"
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
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
      Left            =   2400
      PasswordChar    =   "|"
      TabIndex        =   6
      ToolTipText     =   "Enter your old password"
      Top             =   144
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "PASSWORD"
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
      Left            =   2400
      PasswordChar    =   "|"
      TabIndex        =   5
      ToolTipText     =   "Enter your new password"
      Top             =   550
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "CHANGE"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Changer the old password"
      Top             =   1560
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Goes back to previous page"
      Top             =   1560
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Enter New Password:"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   2115
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
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
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Password:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1830
   End
End
Attribute VB_Name = "changepassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
   
    Me.Hide
    
End Sub

Private Sub cmdOK_Click()

If Text1.Text = "" Then
MsgBox "Enter Your Old Password ", , "Error"
ElseIf Text2 = "" Then
MsgBox "Enter Your New Password ", , "Error"
ElseIf Text3.Text = "" Then
MsgBox "Re-Enter Your New Password ", , "Error"
Else
Data1.RecordSource = "SELECT * FROM ACCOUNT WHERE USERNAME = '" & Text4.Text & "' "
Data1.Refresh
    If Text1.Text = Text5.Text Then
       If Text2.Text = Text3.Text Then
       Dim strfind As String
       Dim intlen As Integer
       strfind = Text2.Text
       intlen = Len(strfind)
          If intlen < 6 Then
          MsgBox "The Password Should Be Atleast Six Character Long ", , "Error"
          Text2.Text = ""
          Text3.Text = ""
          Else
          Text5.Text = Text3.Text
          Data1.Recordset.Edit
          Data1.Recordset.Update
          MsgBox "Password Was Successfully Modified", , ""
          Text1.Text = ""
          Text2.Text = ""
          Text3.Text = ""
          End If
       Else
       MsgBox "The Passwords You Typed Do Not Match.Please Re-Type The New Password In Both Boxes ", , "Error"
       Text2.Text = ""
       Text3.Text = ""
       End If
    Else
    MsgBox "The Old Password Is Incorrect.Please Enter A Valid Password ", , "Error"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    End If
End If
    
End Sub

