VERSION 5.00
Begin VB.Form issualdetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issual of book"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   7710
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CARDDETAILS"
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      RecordSource    =   "ISSUEBOOK"
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      DataField       =   "CALLNO"
      DataSource      =   "Data4"
      Height          =   285
      Left            =   10560
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   8400
      TabIndex        =   17
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DATE"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      DataField       =   "DAYSINTERVAL"
      DataSource      =   "Data3"
      Height          =   285
      Left            =   8400
      TabIndex        =   16
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "COPIES"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   10440
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8400
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ISSUEBOOK"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "CLEAR ALL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Clears all the fields"
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "RENEWALS"
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
      Height          =   285
      Index           =   4
      Left            =   5760
      TabIndex        =   11
      ToolTipText     =   "Enter No times book has been renewed"
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Save to the database"
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "RETURNDATE"
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
      Height          =   285
      Index           =   3
      Left            =   5760
      TabIndex        =   8
      ToolTipText     =   "Enter the date of returning of book"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "ISSUALDATE"
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
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   6
      ToolTipText     =   "Enter date of issual of the book"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "CARDNO"
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
      Height          =   285
      Index           =   0
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Enter Call No of the book"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "CALLNO"
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
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "Enter the card no "
      Top             =   3720
      Width           =   2415
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
      Height          =   495
      Index           =   0
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Goes back to previous page"
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO OF RENEWALS"
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
      Index           =   5
      Left            =   3960
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RETURNING DATE"
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
      Index           =   4
      Left            =   3960
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF ISSUAL"
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
      Index           =   3
      Left            =   3960
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CALL NO"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   3720
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISSUED CARD NO   "
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
      Left            =   3960
      TabIndex        =   3
      Top             =   3240
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " DETAILS OF ISSUAL"
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
      Left            =   3720
      TabIndex        =   1
      Top             =   1440
      Width           =   4830
   End
End
Attribute VB_Name = "issualdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

optionpage.Show
Unload Me

End Sub

Private Sub Command2_Click(Index As Integer)

Text1(2).Text = Date

If issualdetails.Text1(0).Text = "" Or issualdetails.Text1(1).Text = "" Then
MsgBox "Please Type An Option", , "Error"
Else
Data5.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text1(0).Text & "' or CARDNO2='" & Text1(0).Text & "'"
Data5.Refresh
    If Data5.Recordset.RecordCount = 0 Then
    MsgBox "This Card No Does Not Exist,Please Select A Valid Card No", , "Error"
    Text1(0).Text = ""
    Else
    Text6.Text = Text1(0).Text 'New Code
    Data4.RecordSource = "SELECT * FROM ISSUEBOOK WHERE CARDNO = '" & Text6.Text & "' "
    Data4.Refresh
        If Text7.Text = "" Then 'New Code
        intervaldate = CInt(issualdetails.Text5.Text)
        issuedate = CDate(issualdetails.Text1(2).Text)
        issualdetails.Text3.Text = issualdetails.Text1(1).Text
        issualdetails.Data2.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & issualdetails.Text3.Text & "' "
        issualdetails.Data2.Refresh
            If issualdetails.Text4.Text = "" Then
            MsgBox "This Book Does Not Exists.Please Select Another Book", , "Error"
            issualdetails.Text3.Text = ""
            issualdetails.Text4.Text = ""
            issualdetails.Text1(0).Text = ""
            issualdetails.Text1(1).Text = ""
            issualdetails.Text6.Text = ""
            ElseIf issualdetails.Text4.Text <= 1 Then
            MsgBox "No More Copies Of This Book Are Available.Please Select Another Copy", , "Error"
            issualdetails.Text3.Text = ""
            issualdetails.Text4.Text = ""
            issualdetails.Text1(0).Text = ""
            issualdetails.Text1(1).Text = ""
            issualdetails.Text6.Text = ""
            Else
            Data2.Recordset.Edit
            issue = Text4.Text - 1
            Text4.Text = issue
            Data2.Recordset.Update
            Data2.Refresh
            issualdetails.Text1(4).Text = "1"
            issualdetails.Text1(3).Text = DateAdd("d", intervaldate, issuedate)
            Data1.Recordset.AddNew
            issualdetails.Text6.Text = ""
            MsgBox "Record Was Successfull Saved", , ""
            End If
        Else
        MsgBox "This Card Is Already In Use.Please Select Another Card", , "Error" 'New Code
        Text1(0).Text = "" 'New Code
        Text6.Text = "" 'New Code
        Text7.Text = "" 'New Code
        End If
    End If
End If
End Sub

Private Sub Command3_Click()

issualdetails.Text1(0).Text = ""
issualdetails.Text1(1).Text = ""
issualdetails.Text3.Text = ""
issualdetails.Text4.Text = ""

End Sub

