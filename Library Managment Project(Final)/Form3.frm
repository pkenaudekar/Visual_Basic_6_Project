VERSION 5.00
Begin VB.Form deleterecord 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete a book record"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10830
   LinkTopic       =   "Form 3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "DELETE"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Delets an existing record"
      Top             =   5160
      Width           =   1935
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
      Index           =   0
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "Select to delete a record by title"
      Top             =   3840
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
      Left            =   5040
      TabIndex        =   3
      ToolTipText     =   "Select to delete a record by Call No"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
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
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Enter your option here"
      Top             =   2520
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
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
      Index           =   0
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Goes back to previous page"
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DELETE RECORD BY  "
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
      Index           =   0
      Left            =   5040
      TabIndex        =   6
      Top             =   3360
      Width           =   2040
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DELETE A  RECORD"
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
      TabIndex        =   0
      Top             =   1440
      Width           =   4665
   End
End
Attribute VB_Name = "deleterecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

optionpage.Show
Unload Me

End Sub

Private Sub Command2_Click()

If deleterecord.Text1.Text = "" Then
MsgBox "Please Type An Option", , "Error"
Else
deleterecord.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text1.Text & "' "
deleterecord.Data1.Refresh
    If deleterecord.Data1.Recordset.RecordCount = 0 Then
    MsgBox "No Results Found For This Field", , ""
    Text1.Text = ""
    Else
    MsgBox "The Record Was Deleted Successfully", , ""
    Text1.Text = ""
    Data1.Recordset.Delete
    End If
End If

End Sub

Private Sub Command3_Click()

If deleterecord.Text2.Text = "" Then
MsgBox "Please Type An Option", , "Error"
ElseIf Option1(1).Value = False And Option3.Value = False And Option1(1).Value = False Then
MsgBox "Please Select An Option", , "Error"
End If
deleteresult.Text2.Text = ""
If Option1(1) = True Then
deleteresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE TITLE LIKE '*" & Text2.Text & "*' "
deleteresult.Data1.Refresh
    If deleteresult.Data1.Recordset.RecordCount = 0 Then
    MsgBox "No Results Found For This Field", , ""
    Text2.Text = ""
    Option1(1) = False
    Else
    deleteresult.Show
    End If
ElseIf Option3 = True Then
deleteresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE AUTHER  LIKE '*" & Text2.Text & "*'  "
deleteresult.Data1.Refresh
    If deleteresult.Data1.Recordset.RecordCount = 0 Then
    MsgBox "No Results Found For This Field", , ""
    Text2.Text = ""
    Option3 = False
    Else
    deleteresult.Show
    End If
End If

End Sub

Private Sub Command4_Click()

If Text1.Text = "" Then
MsgBox "Please Type An Option", , "Error"
ElseIf Option1(0) = False And Option2 = False Then
MsgBox "Please Select An Option", , "Error"
End If

If Option1(0) = True Then
deleteresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE TITLE = '" & Text1.Text & "' "
deleteresult.Data1.Refresh
    If deleteresult.Data1.Recordset.RecordCount = 0 Then
    MsgBox "This Record Does Not Exists", , "Error"
    Option1(0) = False
    Option2 = False
    Text1.Text = ""
    ElseIf deleteresult.Data1.Recordset.RecordCount = 1 Then
    deleteresult.Data3.RecordSource = "SELECT * FROM ISSUEBOOK WHERE CALLNO= '" & deleteresult.Text3.Text & "' " 'new code
    deleteresult.Data3.Refresh 'new code
        If deleteresult.Data3.Recordset.RecordCount = 1 Then 'new code
        MsgBox "Record In Use,Deletion Denied", , "Error" 'new code
        Text1.Text = "" 'new code
        Option1(0) = False 'new code
        Option2 = False 'new code
        Else 'new code
        deleteresult.Data1.Recordset.Delete
        MsgBox "The Record Was Deleted Successfully", , ""
        Option1(0) = False
        Option2 = False
        Text1.Text = ""
        End If 'new code
    Else
    deleteresult.Text2.Text = ""
    deleteresult.Show
    Unload Me
    End If
End If

If Option2 = True Then
deleteresult.Data3.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text1.Text & "' " 'new code
deleteresult.Data3.Refresh 'new code
    If deleteresult.Data3.Recordset.RecordCount = 1 Then 'new code
    MsgBox "Record In Use,Deletion Denied", , "Error" 'new code
    Text1.Text = "" 'new code
    Option1(0) = False 'new code
    Option2 = False 'new code
    Else 'new code
    deleteresult.Data1.RecordSource = "SELECT * FROM BOOKINFO WHERE CALLNO = '" & Text1.Text & "' "
    deleteresult.Data1.Refresh
        If deleteresult.Data1.Recordset.RecordCount = 0 Then
        MsgBox "This Record Does Not Exists", , "Error"
        Option1(0) = False
        Option2 = False
        Text1.Text = ""
        Else
        deleteresult.Data1.Recordset.Delete
        MsgBox "The Record Was Deleted Successfully", , ""
        Option1(0) = False
        Option2 = False
        Text1.Text = ""
        End If
     End If 'new code
End If

End Sub

