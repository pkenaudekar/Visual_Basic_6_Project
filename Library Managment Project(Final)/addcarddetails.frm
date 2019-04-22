VERSION 5.00
Begin VB.Form addcarddetails 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7635
   ClientLeft      =   105
   ClientTop       =   780
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "addcarddetails.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ISSUEBOOK"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      RecordSource    =   "CARDDETAILS"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CARDDETAILS"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CARDDETAILS"
      Top             =   6120
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   1815
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   1815
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "CARDNO2"
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
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Enter the 2nd card no"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "DEPARTMENT"
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
      ItemData        =   "addcarddetails.frx":3A95
      Left            =   4800
      List            =   "addcarddetails.frx":3AA5
      Sorted          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Section in which department he/she belongs"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "SEMESTER"
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
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "Enter the semester"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAME"
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
      Left            =   4800
      TabIndex        =   2
      ToolTipText     =   "Enter  name of the card holder"
      Top             =   2640
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      DataField       =   "CARDNO1"
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
      Left            =   4800
      TabIndex        =   1
      ToolTipText     =   "Enter the 1st card no"
      Top             =   4080
      Width           =   1455
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
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CARDDETAILS"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARD NO 1"
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
      Left            =   3000
      TabIndex        =   10
      Top             =   4080
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARD NO 2"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   4560
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEMESTER"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   3600
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT "
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
      Left            =   3000
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER CARD DETAILS"
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
      TabIndex        =   0
      Top             =   1080
      Width           =   5265
   End
End
Attribute VB_Name = "addcarddetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1(1).Text = "" Or Text2(1).Text = "" Or Text4.Text = "" Or Text3(2).Text = "" Then
MsgBox "Please Type An Option", , "Error"
ElseIf Combo1(1).Text = "" Then
MsgBox "Please Select A Department", , "Error"
ElseIf Text4.Text = Text3(2).Text Then
MsgBox "A User Cannot Have 2 Cards Of Same No", , "Error"
Text4.Text = ""
Text3(2).Text = ""
Else
Data2.RecordSource = "SELECT * FROM CARDDETAILS WHERE NAME='" & Text1(1).Text & " ' "
Data2.Refresh
Data3.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text4.Text & "' or CARDNO2='" & Text4.Text & "'"
Data3.Refresh
Data4.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text3(2).Text & "' or CARDNO2='" & Text3(2).Text & "'"
Data4.Refresh
    If Data2.Recordset.RecordCount = 1 Then
    MsgBox "This Name Already Exists", , "Error"
    Text1(1).Text = ""
    ElseIf Data3.Recordset.RecordCount = 1 Then
    MsgBox "This Card No-1 Already Exists", , "Error"
    Text4.Text = ""
    ElseIf Data4.Recordset.RecordCount = 1 Then
    MsgBox "This Card No-2 Already Exists", , "Error"
    Text3(2).Text = ""
    Else
    Data1.Recordset.AddNew
    MsgBox "Record Successfully Saved", , ""
    addcarddetails.Text1(1).Text = ""
    addcarddetails.Combo1(1).Text = ""
    addcarddetails.Text2(1).Text = ""
    addcarddetails.Text4.Text = ""
    addcarddetails.Text3(2).Text = ""
    End If
End If
End Sub

Private Sub Command2_Click()
addcarddetails.Text1(1).Text = ""
addcarddetails.Combo1(1).Text = ""
addcarddetails.Text2(1).Text = ""
addcarddetails.Text4.Text = ""
addcarddetails.Text3(2).Text = ""
End Sub

Private Sub Command3_Click()
optionpage.Show
Unload Me
End Sub

