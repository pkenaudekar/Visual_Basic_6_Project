VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form deleteresult 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete result"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "deleteresult.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "deleteresult.frx":3A95
      Height          =   4335
      Left            =   1800
      OleObjectBlob   =   "deleteresult.frx":3AA9
      TabIndex        =   0
      Top             =   1800
      Width           =   8880
   End
   Begin VB.TextBox Text3 
      DataField       =   "CALLNO"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ISSUEBOOK"
      Top             =   600
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
      Height          =   375
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataField       =   "TITLE"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
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
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Goes back to previous page"
      Top             =   6840
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deletes the record"
      Top             =   6360
      Width           =   1500
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
      Left            =   7200
      TabIndex        =   2
      ToolTipText     =   "Enter your option here"
      Top             =   6360
      Width           =   1815
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
      Left            =   5040
      TabIndex        =   5
      Top             =   6360
      Width           =   1965
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
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   4005
   End
End
Attribute VB_Name = "deleteresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

deleterecord.Text1.Text = ""
deleterecord.Option1(0) = False
deleterecord.Option2 = False
deleterecord.Show
Unload Me

End Sub

Private Sub cmdOK_Click()

If Text1.Text = "" Then
MsgBox "Enter A Call No", , "Error"
Else
Data3.RecordSource = "SELECT * FROM ISSUEBOOK WHERE CALLNO ='" & Text1.Text & "' " 'new code
Data3.Refresh 'new code
    If Data3.Recordset.RecordCount = 1 Then 'new code
    MsgBox "Record In Use,Deletion Denied", , "Error" 'new code
    Text1.Text = "" 'new code
    Else 'new code
    Data2.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text1.Text & "' "
    Data2.Refresh
        If Text2.Text = "" Then
        MsgBox "This Record Does Not Exists", , "Error"
        Text1.Text = ""
        Else
        Dim flag As Integer
        flag = 0
        Do While flag = 0
            If Data1.Recordset.Fields("CALLNO") = deleteresult.Text1.Text Then
            Data1.Recordset.Delete
            MsgBox "The Record Was Deleted Successfully", , ""
            deleteresult.Text1.Text = ""
            deleteresult.Text2.Text = ""
            flag = 1
            Data1.Refresh
            Else
            Data1.Recordset.MoveNext
            End If
        Loop
        End If
    End If 'new code
End If

End Sub

