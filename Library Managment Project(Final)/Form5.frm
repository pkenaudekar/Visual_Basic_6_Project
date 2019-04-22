VERSION 5.00
Begin VB.Form startuppage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search page"
   ClientHeight    =   8490
   ClientLeft      =   3510
   ClientTop       =   2670
   ClientWidth     =   11880
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8760
      Top             =   3720
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Close"
      Top             =   240
      Width           =   255
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
      ItemData        =   "Form5.frx":1FA442
      Left            =   4560
      List            =   "Form5.frx":1FA458
      TabIndex        =   8
      Text            =   "All"
      ToolTipText     =   "Select a section here"
      Top             =   4320
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "LIBRARY STAFF LOGIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8160
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "SEARCH"
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
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Search a record in database"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "CALL NO"
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
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "Select to search a record by Call No"
      Top             =   3840
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C00000&
      Caption         =   "AUTHOR"
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
      Left            =   4560
      TabIndex        =   4
      ToolTipText     =   "Select to search a record by author"
      Top             =   3360
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "TITLE OF THE BOOK"
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
      Left            =   4560
      TabIndex        =   3
      ToolTipText     =   "Select to search a record by title"
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "title"
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
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   "Enter your option here"
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IN SECTION"
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
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Top             =   4320
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH  BY     "
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
      Left            =   5400
      TabIndex        =   2
      Top             =   2400
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY BOOK SEARCH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   840
      Width           =   5700
   End
End
Attribute VB_Name = "startuppage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If startuppage.Text1(1).Text = "" Then
MsgBox "Please Type An Option", , "Error"
ElseIf Option1(0).Value = False And Option2.Value = False And Option1(1).Value = False Then
MsgBox "Please Select An Option", , "Error"
ElseIf Combo1.Text = "All" Then
    If Option1(0) = True Then
    searchresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE TITLE LIKE '*" & Text1(1).Text & "*' "
    searchresult.Data1.Refresh
        If searchresult.Data1.Recordset.RecordCount = 0 Then
        MsgBox "No Results Found For This Field", , ""
        Else
        searchresult.Show
        End If
    ElseIf Option2 = True Then
    searchresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE AUTHER  LIKE '*" & Text1(1).Text & "*'  "
    searchresult.Data1.Refresh
        If searchresult.Data1.Recordset.RecordCount = 0 Then
        MsgBox "No Results Found For This Field", , ""
        Else
        searchresult.Show
        End If
    ElseIf Option1(1) = True Then
    searchresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text1(1).Text & "' "
    searchresult.Data1.Refresh
        If searchresult.Data1.Recordset.RecordCount = 0 Then
        MsgBox "No Results Found For This Field", , ""
        Else
        searchresult.Show
        End If
    End If
Else
    If Option1(0) = True Then
        searchresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE TITLE LIKE '*" & Text1(1).Text & "*'AND SECTION ='" & Combo1.Text & "' "
        searchresult.Data1.Refresh
            If searchresult.Data1.Recordset.RecordCount = 0 Then
            MsgBox "No Results Found For This Field", , ""
            Else
            searchresult.Show
            End If
        ElseIf Option2 = True Then
        searchresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE AUTHER  LIKE '*" & Text1(1).Text & "*' AND SECTION ='" & Combo1.Text & "'  "
        searchresult.Data1.Refresh
            If searchresult.Data1.Recordset.RecordCount = 0 Then
            MsgBox "No Results Found For This Field", , ""
            Else
            searchresult.Show
            End If
        ElseIf Option1(1) = True Then
        searchresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text1(1).Text & "'"
        searchresult.Data1.Refresh
            If searchresult.Data1.Recordset.RecordCount = 0 Then
            MsgBox "No Results Found For This Field", , ""
            Else
            searchresult.Show
            End If
        End If
End If

End Sub

Private Sub Command2_Click()


stafflogin.Show
stafflogin.txtUserName.Text = ""
stafflogin.txtPassword.Text = ""
stafflogin.Text1.Text = ""
stafflogin.Text2.Text = ""
stafflogin.Text3.Text = ""
stafflogin.Text4.Text = ""

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Option1(0).Value = True

End Sub

Private Sub Timer1_Timer()
If Option1(1) = True Then
Combo1.Text = "All"
Combo1.Visible = False
Label1(2).Visible = False
Else
Combo1.Visible = True
Label1(2).Visible = True
End If
End Sub

