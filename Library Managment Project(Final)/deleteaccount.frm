VERSION 5.00
Begin VB.Form deleteaccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete an account"
   ClientHeight    =   2175
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "deleteaccount.frx":0000
   ScaleHeight     =   1285.062
   ScaleMode       =   0  'User
   ScaleWidth      =   3718.226
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      DataField       =   "PASSWORD"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   240
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
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ACCOUNT"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
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
      Height          =   360
      Left            =   960
      TabIndex        =   2
      ToolTipText     =   "Enter username here"
      Top             =   840
      Width           =   2415
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Deletes an exsisting account"
      Top             =   1560
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
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
      Height          =   390
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Goes back to previous page"
      Top             =   1560
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER USERNAME"
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
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "deleteaccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()

   Text2.Text = ""
   Me.Hide
   
End Sub

Private Sub cmdOK_Click()

If Text1.Text = "" Then
MsgBox "Enter A Username To Delete", , "Error"
ElseIf Text2.Text = "Y" Then
        If Text1.Text = "administrator" Or Text1.Text = Text4.Text Then
        MsgBox "This Account Cannot Be Deleted", , "Error"
        Text1.Text = ""
        Else
            Data1.RecordSource = "SELECT * FROM ACCOUNT WHERE USERNAME = '" & Text1.Text & "' "
            Data1.Refresh
            If Text3.Text = "" Then
            MsgBox "The Account Does Not Exists", , "Error"
            Text1.Text = ""
            Text3.Text = ""
            Else
            Data1.Recordset.Delete
            MsgBox "The Account Was Deleted Successfully", , ""
            Text1.Text = ""
            Text3.Text = ""
            End If
        End If
End If

End Sub

