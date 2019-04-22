VERSION 5.00
Begin VB.Form newrecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New book record"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      DataField       =   "TITLE"
      DataSource      =   "Data3"
      Height          =   285
      Left            =   8640
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   6720
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   6360
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
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      DataField       =   "AUTHER"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   9720
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   9720
      TabIndex        =   14
      Top             =   2640
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
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
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
      Index           =   1
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Goes back to previous page"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Adds Record to database"
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
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
      Index           =   1
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Clears all the fields above"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox Text4 
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
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Call No of book"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "TITLE"
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
      Left            =   4680
      TabIndex        =   3
      ToolTipText     =   "Enter  name of the book"
      Top             =   2640
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      DataField       =   "AUTHER"
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
      Left            =   4680
      TabIndex        =   2
      ToolTipText     =   "Enter author of the book"
      Top             =   3120
      Width           =   4815
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "SECTION"
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
      Height          =   330
      Index           =   1
      ItemData        =   "Form2.frx":3A95
      Left            =   4680
      List            =   "Form2.frx":3AA8
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Section in which book belongs"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      DataField       =   "COPIES"
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
      Left            =   4680
      TabIndex        =   0
      ToolTipText     =   "Number of copies available"
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER A NEW RECORD"
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
      TabIndex        =   12
      Top             =   1440
      Width           =   5475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TITLE OF BOOK      "
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
      Left            =   2880
      TabIndex        =   9
      Top             =   2640
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AUTHOR  "
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
      Left            =   2880
      TabIndex        =   8
      Top             =   3120
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SECTION"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   3600
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUMBER OF COPIES"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label3 
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
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   4080
      Width           =   750
   End
End
Attribute VB_Name = "newrecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

If newrecord.Text1(1).Text = "" Or newrecord.Text2(1).Text = "" Or newrecord.Text4.Text = "" Or newrecord.Text3(2).Text = "" Or newrecord.Text3(2).Text = "" Or newrecord.Combo1(1).Text = "" Then
MsgBox "Please Type An Option", , "Error"
ElseIf (Not (IsNumeric(newrecord.Text3(2).Text))) Then
MsgBox "Please Enter A Number In Copies Field", , "Error"
Else
Text5.Text = Text1(1).Text
Text7.Text = Text4.Text
Data2.RecordSource = "SELECT * FROM BOOKINFO WHERE TITLE = '" & Text5.Text & "' "
Data2.Refresh
    If Text6.Text = Text2(1).Text Then
    MsgBox "The Record Already Exsists.Please Enter A New Record", , "Error"
    newrecord.Text1(1).Text = ""
    newrecord.Text2(1).Text = ""
    newrecord.Text4.Text = ""
    newrecord.Text3(2).Text = ""
    newrecord.Combo1(1).Text = ""
    Else
    Data3.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text7.Text & "' "
    Data3.Refresh
        If Text8.Text = "" Then
        Data1.Recordset.AddNew
        MsgBox "Record Was Successfull Saved", , ""
        Else
        MsgBox "The Call No Already Exsists.Please Give A New Call No", , "Error"
        Text4.Text = ""
        End If
    End If
End If

End Sub

Private Sub Command2_Click(Index As Integer)

newrecord.Text1(1).Text = ""
newrecord.Text2(1).Text = ""
newrecord.Text4.Text = ""
newrecord.Text3(2).Text = ""
newrecord.Combo1(1).Text = ""

End Sub

Private Sub Command3_Click(Index As Integer)
optionpage.Show
Unload Me

End Sub

Private Sub Command4_Click()

End Sub

