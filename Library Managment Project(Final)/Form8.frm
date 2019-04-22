VERSION 5.00
Begin VB.Form accountsettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Settings"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   7710
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "DELETE AN ACCOUNT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "RENEWAL DATE SETTING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   3735
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
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Goes back to previous page"
      Top             =   5520
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "CREATE AN ACCOUNT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Creats a new login account"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Changes an existing password"
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT AN OPTION"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   4605
   End
End
Attribute VB_Name = "accountsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

changepassword.Show
changepassword.Text1.Text = ""
changepassword.Text2.Text = ""
changepassword.Text3.Text = ""
changepassword.Text4.Text = stafflogin.Text3.Text
changepassword.Text5.Text = ""

End Sub

Private Sub Command2_Click()

newaccount.Data1.Recordset.AddNew
newaccount.Text3.Text = ""
newaccount.Text2.Text = ""
newaccount.Text4.Text = stafflogin.Text4.Text
newaccount.Check1.Value = 0
newaccount.Show

End Sub

Private Sub Command3_Click()

optionpage.Show
Unload Me

End Sub

Private Sub Command4_Click()

datesetting.Show

End Sub

Private Sub Command5_Click()

deleteaccount.Text2.Text = stafflogin.Text4.Text
deleteaccount.Text3.Text = ""
deleteaccount.Text1.Text = ""
deleteaccount.Text4.Text = stafflogin.txtUserName.Text
deleteaccount.Show

End Sub
