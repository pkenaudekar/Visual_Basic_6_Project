VERSION 5.00
Begin VB.Form datesetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Renew Date Interval"
   ClientHeight    =   2130
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "datesetting.frx":0000
   ScaleHeight     =   1258.474
   ScaleMode       =   0  'User
   ScaleWidth      =   3844.983
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DATE"
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "DAYSINTERVAL"
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
      ItemData        =   "datesetting.frx":3A95
      Left            =   1680
      List            =   "datesetting.frx":3AF6
      TabIndex        =   3
      ToolTipText     =   "Select an interval"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "SET"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sets the difference in date"
      Top             =   1440
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Goes back to option"
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SET AN INTERVAL"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   2160
   End
End
Attribute VB_Name = "datesetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()

Combo1.Text = ""
Unload Me

End Sub

Private Sub cmdOK_Click()

Data1.Recordset.Edit
Data1.Recordset.Update
MsgBox "Renewal Date Interval Was Successfull Modified", , ""

End Sub
