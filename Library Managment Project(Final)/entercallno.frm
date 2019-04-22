VERSION 5.00
Begin VB.Form entercallno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter call no"
   ClientHeight    =   2040
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "entercallno.frx":0000
   ScaleHeight     =   1205.299
   ScaleMode       =   0  'User
   ScaleWidth      =   3718.226
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Enter your option here"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Pratik\Desktop\New Folder\Library Managment Project\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "DELETE"
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
      TabIndex        =   0
      ToolTipText     =   "Deletes the record"
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Height          =   390
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Goes back to previous page"
      Top             =   1200
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER CALL NO"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1965
   End
End
Attribute VB_Name = "entercallno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
Unload Me
Unload deleteresult
deleterecord.Text1.Text = ""
deleterecord.Text2.Text = ""
deleterecord.Option1(1) = False
deleterecord.Option3 = False

End Sub

Private Sub cmdOK_Click()

If Text1.Text = "" Then
MsgBox "Enter A Call No", , "Error"
Else
Dim flag As Integer
flag = 0
Do While flag = 0
    If Data1.Recordset.Fields("CALLNO") = entercallno.Text1.Text Then
    Data1.Recordset.Delete
    MsgBox "The Record Was Deleted Successfully", , ""
    entercallno.Text1.Text = ""
    flag = 1
    Data1.Refresh
    Else
    Data1.Recordset.MoveNext
    End If
Loop
End If

End Sub

