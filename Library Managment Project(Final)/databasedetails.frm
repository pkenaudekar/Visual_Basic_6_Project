VERSION 5.00
Begin VB.Form databasedetails 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "databasedetails.frx":0000
   ScaleHeight     =   7695
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "SHOW  LIBRARY CARD DETAILS"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "SHOW ALL BOOK'S ISSUED "
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   3615
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "TODAY'S BOOK RETURN LIST"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "SHOW BOOKS IN LIBRARY"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT A TASK TO PERFORM"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   7080
   End
End
Attribute VB_Name = "databasedetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
booksearchoption.Option1(0) = False
booksearchoption.Option2 = False
booksearchoption.Combo1.Visible = False
booksearchoption.Combo1.Text = ""
booksearchoption.Show
End Sub

Private Sub Command2_Click()
bookreturnlist.Data1.RecordSource = "SELECT * FROM ISSUEBOOK WHERE RETURNDATE <= Date() "
bookreturnlist.Data1.Refresh
If bookreturnlist.Data1.Recordset.RecordCount = 0 Then
MsgBox "There Are No Book's To Be Returned Today", , ""
Else
bookreturnlist.Text1.Text = bookreturnlist.Data1.Recordset.RecordCount
bookreturnlist.Show
Unload Me
End If
End Sub

Private Sub Command3_Click()
optionpage.Show
Unload Me
End Sub

Private Sub Command4_Click()
bookreturnlist.Data1.RecordSource = "SELECT * FROM ISSUEBOOK WHERE CARDNO LIKE '*'"
bookreturnlist.Data1.Refresh
If bookreturnlist.Data1.Recordset.RecordCount = 0 Then
MsgBox "There Are No Book's Issued", , ""
Else
bookreturnlist.Text1.Text = bookreturnlist.Data1.Recordset.RecordCount
bookreturnlist.Show
Unload Me
End If
End Sub

Private Sub Command5_Click()
librarycardoption.Option1(0) = False
librarycardoption.Option2 = False
librarycardoption.Text1.Visible = False
librarycardoption.Text1.Text = ""
librarycardoption.Show
End Sub
