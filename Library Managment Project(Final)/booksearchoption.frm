VERSION 5.00
Begin VB.Form booksearchoption 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "booksearchoption.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4200
      Top             =   960
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
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
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "OK"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "booksearchoption.frx":3A95
      Left            =   720
      List            =   "booksearchoption.frx":3AA8
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "SHOW ALL BOOKS"
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
      Height          =   345
      Index           =   0
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Select to search a record by title"
      Top             =   720
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C00000&
      Caption         =   "SHOW BOOK IN A SECTION"
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
      Height          =   345
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Select to search a record by author"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT AN OPTION    "
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
      Index           =   1
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   2595
   End
End
Attribute VB_Name = "booksearchoption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1(0) = False And Option2 = False Then
MsgBox "Please Select An Option", , "Error"
ElseIf Option2 = True And Combo1.Text = "" Then
MsgBox "Please Select A Section", , "Error"
ElseIf Option1(0) = True Then
booksindatabase.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE SECTION LIKE '*'"
booksindatabase.Data1.Refresh
    If booksindatabase.Data1.Recordset.RecordCount = 0 Then
    MsgBox "There Are No Records In Database", , ""
    Else
    booksindatabase.Text1.Text = booksindatabase.Data1.Recordset.RecordCount
    booksindatabase.Show
    Unload Me
    Unload databasedetails
    End If
ElseIf Option2 = True Then
booksindatabase.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE SECTION ='" & Combo1.Text & "'"
booksindatabase.Data1.Refresh
    If booksindatabase.Data1.Recordset.RecordCount = 0 Then
    MsgBox "There Are No Records In This Section In Database", , ""
    Else
    booksindatabase.Text1.Text = booksindatabase.Data1.Recordset.RecordCount
    booksindatabase.Show
    Unload Me
    Unload databasedetails
    End If
End If
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Timer1_Timer()
If Option1(0) = True Then
Combo1.Visible = False
Combo1.Text = ""
ElseIf Option2 = True Then
Combo1.Visible = True
End If
End Sub
