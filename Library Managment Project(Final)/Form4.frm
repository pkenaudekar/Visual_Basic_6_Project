VERSION 5.00
Begin VB.Form modifyrecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify a record"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11625
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   8565
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   7080
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ISSUEBOOK"
      Top             =   7920
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
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text6 
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
      Left            =   3600
      TabIndex        =   19
      ToolTipText     =   "Enter your option here"
      Top             =   1440
      Width           =   5175
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   9000
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "CLEAR"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Clears all the fields"
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BOOKINFO"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "MODIFY"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Modifys an existint record"
      Top             =   7080
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "TITLE OF THE BOOK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      ToolTipText     =   "Select to search by title"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C00000&
      Caption         =   "CALL NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      ToolTipText     =   "Select to search by Call No"
      Top             =   2760
      Width           =   2055
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
      Left            =   5040
      TabIndex        =   8
      ToolTipText     =   "Number of copies available"
      Top             =   6000
      Width           =   975
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
      ItemData        =   "Form4.frx":1FA442
      Left            =   5040
      List            =   "Form4.frx":1FA455
      Sorted          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Section in which book belongs"
      Top             =   5040
      Width           =   2535
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
      Left            =   5040
      TabIndex        =   6
      ToolTipText     =   "Enter author of the book"
      Top             =   4560
      Width           =   4815
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
      Left            =   5040
      TabIndex        =   5
      ToolTipText     =   "Enter title of the book"
      Top             =   4080
      Width           =   4815
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
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "Call No of book"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Search a record"
      Top             =   3240
      Width           =   1095
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Goes back to previous page"
      Top             =   7080
      Width           =   1575
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
      Left            =   3240
      TabIndex        =   13
      Top             =   5520
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
      Left            =   3240
      TabIndex        =   12
      Top             =   6000
      Width           =   1695
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
      Left            =   3240
      TabIndex        =   11
      Top             =   5040
      Width           =   750
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
      Left            =   3240
      TabIndex        =   10
      Top             =   4560
      Width           =   810
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
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Top             =   4080
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH RECORD BY  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   5160
      TabIndex        =   2
      Top             =   1920
      Width           =   2100
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MODIFY A RECORD"
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
      Top             =   600
      Width           =   4470
   End
End
Attribute VB_Name = "modifyrecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If modifyrecord.Text6.Text = "" Then
MsgBox "Please Type An Option", , "Error"
ElseIf Option1.Value = False And Option2.Value = False Then
MsgBox "Please Select An Option", , "Error"
Else
Text5.Text = Text6.Text
    If Option1.Value = True Then
    modifyresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE TITLE = '" & Text5.Text & "' "
    modifyresult.Data1.Refresh
        If modifyresult.Data1.Recordset.RecordCount > 1 Then
        modifyresult.Text1.Text = ""
        modifyresult.Text2.Text = ""
        modifyrecord.Text6.Text = "" 'new code
        modifyrecord.Option1.Value = False
        modifyrecord.Option2.Value = False
        modifyresult.Show
        Unload Me
        ElseIf modifyresult.Data1.Recordset.RecordCount = 0 Then
        MsgBox "This Record Does Not Exists", , "Error"
        modifyrecord.Text6.Text = "" 'new code
        modifyrecord.Text5.Text = "" 'new code
        Option1.Value = False
        Option2.Value = False
        ElseIf modifyresult.Data1.Recordset.RecordCount = 1 Then
        Text7.Text = modifyresult.Data1.Recordset.Fields("CALLNO")
        modifyrecord.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE TITLE = '" & Text5.Text & "' "
        modifyrecord.Data1.Refresh
        modifyrecord.Text6.Text = "" 'new code
        Option1.Value = False
        Option2.Value = False
        End If
    End If
    If Option2.Value = True Then
    Text7.Text = Text6.Text
    Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text5.Text & "' "
    Data1.Refresh
    modifyrecord.Text6.Text = ""
    Option1.Value = False
    Option2.Value = False
        If Text1(1).Text = "" Then
        MsgBox "This Record Does Not Exists", , "Error"
        modifyrecord.Text5.Text = ""
        Option1.Value = False
        Option2.Value = False
        End If
    End If
End If

End Sub

Private Sub Command2_Click()

If modifyrecord.Text1(1).Text = "" And modifyrecord.Text2(1).Text = "" And modifyrecord.Combo1(1).Text = "" And modifyrecord.Text4.Text = "" And modifyrecord.Text3(2).Text = "" Then
MsgBox "Please Select A Record To Modify ", , "Error"
ElseIf modifyrecord.Text1(1).Text = "" Or modifyrecord.Text2(1).Text = "" Or modifyrecord.Combo1(1).Text = "" Or modifyrecord.Text4.Text = "" Or modifyrecord.Text3(2).Text = "" Then
MsgBox "Please Fill All The Fields ", , "Error"
ElseIf (Not (IsNumeric(modifyrecord.Text3(2).Text))) Then
MsgBox "Please Enter A Number In The Number Of Copies Field", , "Error"
modifyrecord.Text3(2).Text = ""
Else
Data3.RecordSource = "SELECT * FROM ISSUEBOOK WHERE CALLNO ='" & Text7.Text & "' "
Data3.Refresh
    If Data3.Recordset.RecordCount = 1 Then
    MsgBox "This Record Is In Use,Modification Denied", , "Error"
    modifyrecord.Text5.Text = ""
    modifyrecord.Text6.Text = ""
    modifyrecord.Text1(1).Text = ""
    modifyrecord.Text2(1).Text = ""
    modifyrecord.Combo1(1).Text = ""
    modifyrecord.Text4.Text = ""
    modifyrecord.Text3(2).Text = ""
    Text7.Text = ""
    Else
        If Text4.Text = Text7.Text Then
        Data1.Recordset.Edit
        Data1.Recordset.Update
        MsgBox "Record Was Successfull Modified", , ""
        modifyrecord.Text5.Text = ""
        modifyrecord.Text6.Text = ""
        modifyrecord.Text1(1).Text = ""
        modifyrecord.Text2(1).Text = ""
        modifyrecord.Combo1(1).Text = ""
        modifyrecord.Text4.Text = ""
        modifyrecord.Text3(2).Text = ""
        Text7.Text = ""
        Option1 = False
        Option2 = False
        Else
        Data2.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO ='" & Text4.Text & "' "
        Data2.Refresh
            If Data2.Recordset.RecordCount = 1 Then
            MsgBox "This Card No Already Exists", , "Error"
            Text4.Text = ""
            Else
            Data1.Recordset.Edit
            Data1.Recordset.Update '
            MsgBox "Record Was Successfull Modified", , ""
            modifyrecord.Text5.Text = ""
            modifyrecord.Text6.Text = ""
            modifyrecord.Text1(1).Text = ""
            modifyrecord.Text2(1).Text = ""
            modifyrecord.Combo1(1).Text = ""
            modifyrecord.Text4.Text = ""
            modifyrecord.Text3(2).Text = ""
            Text7.Text = ""
            Option1 = False
            Option2 = False
            End If
        End If
    End If
End If

End Sub

Private Sub Command3_Click(Index As Integer)

modifyrecord.Text5.Text = ""
modifyrecord.Text6.Text = ""
modifyrecord.Text1(1).Text = ""
modifyrecord.Text2(1).Text = ""
modifyrecord.Combo1(1).Text = ""
modifyrecord.Text4.Text = ""
modifyrecord.Text3(2).Text = ""
Text7.Text = ""
Option1 = False
Option2 = False
optionpage.Show
Unload Me

End Sub

Private Sub Command4_Click()

modifyrecord.Text5.Text = ""
modifyrecord.Text6.Text = ""
modifyrecord.Text1(1).Text = ""
modifyrecord.Text2(1).Text = ""
modifyrecord.Combo1(1).Text = ""
modifyrecord.Text4.Text = ""
modifyrecord.Text3(2).Text = ""
Text7.Text = ""
Option1 = False
Option2 = False

End Sub

