VERSION 5.00
Begin VB.Form optionpage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option page"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "library332.frx":0000
   ScaleHeight     =   8610
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      BackColor       =   &H008080FF&
      Caption         =   "ADD/DELETE/MODIFY CARD DETAILS"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Select to add/delete/modify card details"
      Top             =   5520
      Width           =   4455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   "RECORD DETAILS"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   " SETTINGS"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      Caption         =   "LOGOUT"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Logs out of database"
      Top             =   7680
      Width           =   4455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "RETURN A BOOK"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Return/renew a book period"
      Top             =   4800
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "ISSUE A BOOK"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Keep a record of issued books"
      Top             =   4080
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   " MODIFY A RECORD"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Edits an existing record"
      Top             =   3360
      Width           =   4455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "DELETE A  RECORD"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Deletes an exitsting record from database"
      Top             =   2640
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "ADD A NEW RECORD"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Adds a new record in database"
      Top             =   1920
      Width           =   4455
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
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   7080
   End
End
Attribute VB_Name = "optionpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

newrecord.Data1.Recordset.AddNew
newrecord.Text6.Text = ""
newrecord.Text8.Text = ""
newrecord.Show
optionpage.Visible = False
Unload Me

End Sub

Private Sub Command2_Click()

If stafflogin.Text4.Text = "Y" Then
accountsettings.Command2.Enabled = True
accountsettings.Command5.Enabled = True
accountsettings.Command4.Enabled = True

Else
accountsettings.Command2.Enabled = False
accountsettings.Command5.Enabled = False
accountsettings.Command4.Enabled = False

End If
accountsettings.Show
Unload Me

End Sub

Private Sub Command3_Click(Index As Integer)

modifyrecord.Show
optionpage.Visible = False
Unload Me
modifyrecord.Text5.Text = ""
modifyrecord.Text6.Text = ""
modifyrecord.Text1(1).Text = ""
modifyrecord.Text2(1).Text = ""
modifyrecord.Combo1(1).Text = ""
modifyrecord.Text4.Text = ""
modifyrecord.Text7.Text = ""
modifyrecord.Text3(2).Text = ""

End Sub

Private Sub Command4_Click(Index As Integer)

issualdetails.Data1.Recordset.AddNew
issualdetails.Text3.Text = ""
issualdetails.Text4.Text = ""
issualdetails.Text7.Text = ""
issualdetails.Show
optionpage.Visible = False
Unload Me

End Sub

Private Sub Command5_Click(Index As Integer)

detailsofbook.Show
optionpage.Visible = False
Unload Me
detailsofbook.Text1(1).Text = ""
detailsofbook.Text1(0).Text = ""
detailsofbook.Text1(2).Text = ""
detailsofbook.Text1(3).Text = ""
detailsofbook.Text1(4).Text = ""
detailsofbook.Text3.Text = ""
detailsofbook.Text7.Text = ""
detailsofbook.Text6.BackColor = &H80000005

End Sub

Private Sub Command6_Click(Index As Integer)

startuppage.Show
Unload Me

End Sub

Private Sub Command7_Click()

deleterecord.Show
optionpage.Visible = False
Unload Me
deleterecord.Text1.Text = ""

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Command8_Click()
databasedetails.Show
Unload Me
End Sub

Private Sub Command9_Click()
carddetails.Option1(0) = False
carddetails.Option2 = False
carddetails.Option1(1) = False
carddetails.Text1.Text = ""
carddetails.Text1.Visible = False
carddetails.Text2.Text = ""
carddetails.Show
End Sub

