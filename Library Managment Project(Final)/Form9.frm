VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form modifyresult 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify result"
   ClientHeight    =   8505
   ClientLeft      =   -360
   ClientTop       =   630
   ClientWidth     =   11985
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   8505
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form9.frx":3A95
      Height          =   4335
      Left            =   1560
      OleObjectBlob   =   "Form9.frx":3AA9
      TabIndex        =   0
      Top             =   1800
      Width           =   8880
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   480
      Visible         =   0   'False
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
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Enter your option here"
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "MODIFY"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Modifys an existint record"
      Top             =   6480
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Goes back to previous page"
      Top             =   6960
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      DataField       =   "TITLE"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   3840
      TabIndex        =   6
      Top             =   960
      Width           =   4005
   End
   Begin VB.Label Label5 
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
      Left            =   4800
      TabIndex        =   5
      Top             =   6490
      Width           =   1965
   End
End
Attribute VB_Name = "modifyresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
modifyrecord.Text6.Text = ""
modifyrecord.Text1(1).Text = ""
modifyrecord.Text2(1).Text = ""
modifyrecord.Combo1(1).Text = ""
modifyrecord.Text4.Text = ""
modifyrecord.Text3(2).Text = ""
modifyrecord.Option1 = False
modifyrecord.Option2 = False
modifyrecord.Show
Unload Me

End Sub

Private Sub cmdOK_Click()
modifyrecord.Text7.Text = Text1.Text
If Text1.Text = "" Then
MsgBox "Enter A Call No", , "Error"
Else
modifyresult.Data2.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text1.Text & "' "
modifyresult.Data2.Refresh
    If Text2.Text = "" Then
    MsgBox "This Record Does Not Exists", , "Error"
    Text1.Text = ""
    Else
    modifyrecord.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text1.Text & "' "
    modifyrecord.Data1.Refresh
    modifyrecord.Text6.Text = ""
    modifyrecord.Show
    Unload Me
    End If
End If

End Sub
