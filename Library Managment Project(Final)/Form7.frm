VERSION 5.00
Begin VB.Form detailsofbook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return book"
   ClientHeight    =   7725
   ClientLeft      =   105
   ClientTop       =   780
   ClientWidth     =   10830
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
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
      Left            =   5640
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text6 
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
      Left            =   5640
      TabIndex        =   20
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Text5 
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
      Left            =   5640
      TabIndex        =   19
      Top             =   4440
      Width           =   2415
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
      RecordSource    =   "DATE"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "DAYSINTERVAL"
      DataSource      =   "Data3"
      Height          =   285
      Left            =   10920
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "COPIES"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   10920
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8280
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   7080
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
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ISSUEBOOK"
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "RENEW"
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
      TabIndex        =   15
      ToolTipText     =   "Renews a book"
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "RETURN"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Deletes the record from the database"
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "RENEWALS"
      DataSource      =   "Data1"
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
      Index           =   4
      Left            =   5640
      TabIndex        =   11
      ToolTipText     =   "No of renewals performed"
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataField       =   "CALLNO"
      DataSource      =   "Data1"
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
      Index           =   0
      Left            =   5640
      TabIndex        =   7
      ToolTipText     =   "Call No of book issued"
      Top             =   3960
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
      Left            =   8280
      TabIndex        =   6
      ToolTipText     =   "Date on which book was issued"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
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
      Left            =   8280
      TabIndex        =   5
      ToolTipText     =   "Date on which book is to be returned"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "title"
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
      Height          =   315
      Index           =   1
      Left            =   5640
      TabIndex        =   2
      ToolTipText     =   "Enter the Card No who's details are to be searched"
      Top             =   2400
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Goes back to previous page"
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIFFERENCE IN DAYS"
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
      Left            =   3720
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOOK DETAILS"
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
      Index           =   5
      Left            =   5280
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
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
      Index           =   0
      Left            =   3720
      TabIndex        =   12
      Top             =   5400
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " CALL NO"
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
      Left            =   3720
      TabIndex        =   10
      Top             =   3960
      Width           =   795
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
      Left            =   3720
      TabIndex        =   9
      Top             =   4440
      Width           =   1395
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
      Left            =   3720
      TabIndex        =   8
      Top             =   4920
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER ISSUED CARD NO   "
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
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   2190
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " DETAILS OF BOOK RETURN"
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
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   6570
   End
End
Attribute VB_Name = "detailsofbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
detailsofbook.Text1(0).Text = ""
detailsofbook.Text1(2).Text = ""
detailsofbook.Text1(3).Text = ""
detailsofbook.Text1(4).Text = ""
optionpage.Show
Unload Me

End Sub

Private Sub Command2_Click()

If detailsofbook.Text1(1).Text = "" Then
MsgBox "Please Type An Option", , "Error"
Else
Text6.BackColor = &H80000005
Label2.Visible = False
Text7.Visible = False
Data1.RecordSource = "SELECT * FROM ISSUEBOOK WHERE CARDNO = '" & Text1(1).Text & "' "
Data1.Refresh
    If Text1(0).Text = "" Then
    detailsofbook.Text1(0).Text = ""
    detailsofbook.Text1(4).Text = ""
    detailsofbook.Text5.Text = ""
    detailsofbook.Text6.Text = ""
    detailsofbook.Text7.Text = ""
    MsgBox "No Book Is Issued On This Card No", , ""
    Else
    a = CDate(Text1(2).Text)
    b = CDate(Text1(3).Text)
    Text5.Text = Format(a, "medium date")
    Text6.Text = Format(b, "medium date")
    Dim c As Date
    Dim d As Date
    c = Date
    d = CDate(Text6.Text)
        If c > d Then
        Text6.BackColor = &HFF&
        Label2.Visible = True
        Text7.Visible = True
        Text7.Text = DateDiff("d", d, c)
        End If
    End If
End If

End Sub

Private Sub Command3_Click(Index As Integer)

If Text1(0).Text = "" And Text5.Text = "" And Text6.Text = "" And Text1(4).Text = "" And Text7.Text = "" Then
MsgBox "Please Select A Card No", , "Error"
Else
detailsofbook.Text2.Text = detailsofbook.Text1(0).Text
detailsofbook.Data2.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & detailsofbook.Text2.Text & "' " 'New code
detailsofbook.Data2.Refresh 'New code
detailsofbook.Data2.Recordset.Edit
Copies = detailsofbook.Text3.Text + 1
detailsofbook.Text3.Text = Copies
detailsofbook.Data2.Recordset.Update
Data1.Recordset.Delete
detailsofbook.Text1(0).Text = ""
detailsofbook.Text1(2).Text = ""
detailsofbook.Text1(3).Text = ""
detailsofbook.Text1(4).Text = ""
detailsofbook.Text5.Text = ""
detailsofbook.Text6.Text = ""
detailsofbook.Text7.Text = ""
Label2.Visible = False
Text7.Visible = False
Text6.BackColor = &H80000005
MsgBox "The Record Was Deleted Successfully", , ""
End If

End Sub

Private Sub Command4_Click()

If Text1(0).Text = "" And Text5.Text = "" And Text6.Text = "" And Text1(4).Text = "" And Text7.Text = "" Then
MsgBox "Please Select A Card No", , "Error"
Else
Data1.Recordset.Edit
If Text6.BackColor = &HFF& Then
Dim d As Date
d = Date
c = CInt(Text4.Text)
Text1(3).Text = DateAdd("d", c, d)
renew = Text1(4).Text + 1
Text1(4).Text = renew
Text6.Text = Format(Text1(3).Text, "medium date")
Label2.Visible = False
Text7.Visible = False
Text6.BackColor = &H80000005
Else
a = CDate(Text1(3).Text)
b = CInt(Text4.Text)
Text1(3).Text = DateAdd("d", b, a)
renew = Text1(4).Text + 1
Text1(4).Text = renew
Text6.Text = Format(Text1(3).Text, "medium date")
Label2.Visible = False
Text7.Visible = False
Text6.BackColor = &H80000005
End If
Data1.Recordset.Update
Data1.Refresh
End If
End Sub

