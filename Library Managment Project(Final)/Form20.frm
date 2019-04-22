VERSION 5.00
Begin VB.Form Form20 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "date"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3600
      Width           =   2295
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Date
Dim b As Date
'Text1.Text = Date
a = CDate(Text1.Text)
'a = Date
b = CDate(Text2.Text)
'If b > a Then
'Text2.BackColor = &HFF&
Text3.Text = DateDiff("d", b, a)
'End If
End Sub

Private Sub Command2_Click()
Text2.BackColor = &H8000000F
End Sub
