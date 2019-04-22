VERSION 5.00
Begin VB.Form carddetails 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "carddetails.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4080
      Top             =   120
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
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
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
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
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
      Height          =   285
      Left            =   720
      TabIndex        =   3
      ToolTipText     =   "Enter card no here"
      Top             =   2040
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "ADD CARD DETAILS"
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
      Height          =   345
      Index           =   0
      Left            =   720
      TabIndex        =   2
      ToolTipText     =   "Select to add a new library card"
      Top             =   600
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C00000&
      Caption         =   "DELETE CARD DETAILS"
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
      Height          =   345
      Left            =   720
      TabIndex        =   1
      ToolTipText     =   "Select to delete an exsisting card"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "MODIFY CARD DETAILS"
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
      Height          =   345
      Index           =   1
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Select to modify details of a card"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT AN OPTION   "
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
      TabIndex        =   6
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "carddetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1(0) = False And Option1(1) = False And Option2 = False Then
MsgBox "Please Select An Option", , "Error"
ElseIf Option2 = True And Text1.Text = "" Then
MsgBox "Please Enter A Card No", , "Error"
ElseIf Option1(1) = True And Text1.Text = "" Then
MsgBox "Please Enter A Card No", , "Error"
ElseIf Option1(0) = True Then
addcarddetails.Data1.Recordset.AddNew
addcarddetails.Text1(1).Text = ""
addcarddetails.Combo1(1).Text = ""
addcarddetails.Text2(1).Text = ""
addcarddetails.Text4.Text = ""
addcarddetails.Text3(2).Text = ""
addcarddetails.Show
Unload Me
Unload optionpage
ElseIf Option2 = True Then
addcarddetails.Data1.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text1.Text & "' or CARDNO2='" & Text1.Text & "'"
addcarddetails.Data1.Refresh
    If addcarddetails.Data1.Recordset.RecordCount = 0 Then
    MsgBox "This Card No Does Not Exsist", , "Error"
    carddetails.Text1.Text = ""
    ElseIf addcarddetails.Data1.Recordset.RecordCount = 1 Then
        If addcarddetails.Data1.Recordset.Fields("CARDNO1") = Text1.Text Then
        Text2.Text = addcarddetails.Data1.Recordset.Fields("CARDNO2")
        ElseIf addcarddetails.Data1.Recordset.Fields("CARDNO2") = Text1.Text Then
        Text2.Text = addcarddetails.Data1.Recordset.Fields("CARDNO1")
        End If
    addcarddetails.Data5.RecordSource = "SELECT * FROM ISSUEBOOK WHERE CARDNO='" & Text1.Text & "'or CARDNO='" & Text2.Text & "' "
    addcarddetails.Data5.Refresh
        If addcarddetails.Data5.Recordset.RecordCount = 1 Then
        MsgBox "Cards Of This User In Use,Deletion Denied", , "Error"
        Text1.Text = ""
        Text2.Text = ""
        Else
        addcarddetails.Data1.Recordset.Delete
        MsgBox "Card Details Deleted Successfully", , ""
        carddetails.Text1.Text = ""
        Text2.Text = ""
        End If
    End If
ElseIf Option1(1) = True Then
modifycarddetails.Data1.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text1.Text & "' or CARDNO2='" & Text1.Text & "'"
modifycarddetails.Data1.Refresh
    If modifycarddetails.Data1.Recordset.RecordCount = 0 Then
    MsgBox "This Card No Does Not Exsist", , "Error"
    carddetails.Text1.Text = ""
    ElseIf modifycarddetails.Data1.Recordset.RecordCount = 1 Then
        If modifycarddetails.Data1.Recordset.Fields("CARDNO1") = Text1.Text Then
        modifycarddetails.Text9.Text = modifycarddetails.Data1.Recordset.Fields("CARDNO2")
        modifycarddetails.Text5.Text = modifycarddetails.Data1.Recordset.Fields("NAME")
        modifycarddetails.Text6.Text = modifycarddetails.Data1.Recordset.Fields("CARDNO1")
        modifycarddetails.Text7.Text = modifycarddetails.Data1.Recordset.Fields("CARDNO2")
        ElseIf modifycarddetails.Data1.Recordset.Fields("CARDNO2") = Text1.Text Then
        modifycarddetails.Text9.Text = modifycarddetails.Data1.Recordset.Fields("CARDNO1")
        modifycarddetails.Text5.Text = modifycarddetails.Data1.Recordset.Fields("NAME")
        modifycarddetails.Text6.Text = modifycarddetails.Data1.Recordset.Fields("CARDNO1")
        modifycarddetails.Text7.Text = modifycarddetails.Data1.Recordset.Fields("CARDNO2")
        End If
    modifycarddetails.Text8.Text = Text1.Text
    modifycarddetails.Show
    Unload Me
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If Option1(0) = True Then
Text1.Visible = False
Text1.Text = ""
ElseIf Option2 = True Then
Text1.Visible = True
ElseIf Option1(1) = True Then
Text1.Visible = True
End If
End Sub
