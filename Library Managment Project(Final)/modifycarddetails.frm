VERSION 5.00
Begin VB.Form modifycarddetails 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "modifycarddetails.frx":0000
   ScaleHeight     =   4755
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CARDDETAILS"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3840
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3600
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3600
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6960
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CARDDETAILS"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CARDDETAILS"
      Top             =   3360
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
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ISSUEBOOK"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "DEPARTMENT"
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
      ItemData        =   "modifycarddetails.frx":1FA442
      Left            =   2040
      List            =   "modifycarddetails.frx":1FA452
      Sorted          =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "Section in which department he/she belongs"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Library Managment Project(Final)\LIBRARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CARDDETAILS"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      DataField       =   "CARDNO2"
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
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   "Enter the 2nd card no"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "SEMESTER"
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
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Enter the semester"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAME"
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
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Enter  name of the card holder"
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      DataField       =   "CARDNO1"
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
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "Enter the 1st card no"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARD NO 1"
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
      Left            =   720
      TabIndex        =   9
      Top             =   2520
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARD NO 2"
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
      Left            =   720
      TabIndex        =   8
      Top             =   3000
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEMESTER"
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
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT "
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
      Left            =   720
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
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
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRARY CARD DETAILS  "
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
      Index           =   0
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   3045
   End
End
Attribute VB_Name = "modifycarddetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1(1).Text = "" Then
MsgBox "Please Enter A Name", , "Error"
ElseIf Combo1(1).Text = "" Then
MsgBox "Please Select A Department", , "Error"
ElseIf Text2(1).Text = "" Then
MsgBox "Please Enter Semester", , "Error"
ElseIf Text4.Text = "" Then
MsgBox "Please Enter Card No 1", , "Error"
ElseIf Text3(2).Text = "" Then
MsgBox "Please Enter Card No 2", , "Error"
Else
Data2.RecordSource = "SELECT * FROM ISSUEBOOK WHERE CARDNO='" & Text8.Text & "' or CARDNO='" & Text9.Text & "'"
Data2.Refresh
    If Data2.Recordset.RecordCount = 1 Then
    MsgBox "Modification Denied,Cards In Use", , "Error"
    Text1(1).Text = ""
    Combo1(1).Text = ""
    Text2(1).Text = ""
    Text4.Text = ""
    Text3(2).Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Me.Hide
    Else
        If Text1(1).Text = Text5.Text Then
            If Text4.Text = Text6.Text Then
                If Text3(2).Text = Text7.Text Then
                    If Text4.Text = Text3(2).Text Then
                    MsgBox "User Cannot Have Two Card Of Same No", , "Error"
                    Text3(2).Text = ""
                    Text4.Text = ""
                    Else
                    Data1.Recordset.Edit
                    Data1.Recordset.Update
                    MsgBox "Record Was Successfull Modified", , ""
                    Text1(1).Text = ""
                    Combo1(1).Text = ""
                    Text2(1).Text = ""
                    Text4.Text = ""
                    Text3(2).Text = ""
                    Text5.Text = ""
                    Text6.Text = ""
                    Text7.Text = ""
                    Text8.Text = ""
                    Text9.Text = ""
                    Me.Hide
                    End If
                Else
                Data5.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text3(2).Text & "' or CARDNO2='" & Text3(2).Text & "' "
                Data5.Refresh
                    If Data5.Recordset.RecordCount = 1 Then
                    MsgBox "This Card No Already Exsists", , "Error"
                    Text3(2).Text = ""
                    Else
                        If Text4.Text = Text3(2).Text Then
                        MsgBox "User Cannot Have Two Card Of Same No", , "Error"
                        Text3(2).Text = ""
                        Text4.Text = ""
                        Else
                        Data1.Recordset.Edit
                        Data1.Recordset.Update
                        MsgBox "Record Was Successfull Modified", , ""
                        Text1(1).Text = ""
                        Combo1(1).Text = ""
                        Text2(1).Text = ""
                        Text4.Text = ""
                        Text3(2).Text = ""
                        Text5.Text = ""
                        Text6.Text = ""
                        Text7.Text = ""
                        Text8.Text = ""
                        Text9.Text = ""
                        Me.Hide
                        End If
                    End If
                End If
            Else
            Data4.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text4.Text & "'or CARDNO2='" & Text4.Text & "' "
            Data4.Refresh
                If Data4.Recordset.RecordCount = 1 Then
                MsgBox "This Card No Already Exsists", , "Error"
                Text4.Text = ""
                Else
                    If Text3(2).Text = Text7.Text Then
                        If Text4.Text = Text3(2).Text Then
                        MsgBox "User Cannot Have Two Card Of Same No", , "Error"
                        Text3(2).Text = ""
                        Text4.Text = ""
                        Else
                        Data1.Recordset.Edit
                        Data1.Recordset.Update
                        MsgBox "Record Was Successfull Modified", , ""
                        Text1(1).Text = ""
                        Combo1(1).Text = ""
                        Text2(1).Text = ""
                        Text4.Text = ""
                        Text3(2).Text = ""
                        Text5.Text = ""
                        Text6.Text = ""
                        Text7.Text = ""
                        Text8.Text = ""
                        Text9.Text = ""
                        Me.Hide
                        End If
                    Else
                    Data5.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text3(2).Text & "' or CARDNO2='" & Text3(2).Text & "' "
                    Data5.Refresh
                        If Data5.Recordset.RecordCount = 1 Then
                        MsgBox "This Card No Already Exsists", , "Error"
                        Text3(2).Text = ""
                        Else
                            If Text4.Text = Text3(2).Text Then
                            MsgBox "User Cannot Have Two Card Of Same No", , "Error"
                            Text3(2).Text = ""
                            Text4.Text = ""
                            Else
                            Data1.Recordset.Edit
                            Data1.Recordset.Update
                            MsgBox "Record Was Successfull Modified", , ""
                            Text1(1).Text = ""
                            Combo1(1).Text = ""
                            Text2(1).Text = ""
                            Text4.Text = ""
                            Text3(2).Text = ""
                            Text5.Text = ""
                            Text6.Text = ""
                            Text7.Text = ""
                            Text8.Text = ""
                            Text9.Text = ""
                            Me.Hide
                            End If
                        End If
                    End If
                End If
            End If
        Else
        Data3.RecordSource = "SELECT * FROM CARDDETAILS WHERE NAME='" & Text1(1).Text & "' "
        Data3.Refresh
            If Data3.Recordset.RecordCount = 1 Then
            MsgBox "This Name Already Exsists", , "Error"
            Text1(1).Text = ""
            Else
                If Text4.Text = Text6.Text Then
                    If Text3(2).Text = Text7.Text Then
                        If Text4.Text = Text3(2).Text Then
                        MsgBox "User Cannot Have Two Card Of Same No", , "Error"
                        Text3(2).Text = ""
                        Text4.Text = ""
                        Else
                        Data1.Recordset.Edit
                        Data1.Recordset.Update
                        MsgBox "Record Was Successfull Modified", , ""
                        Text1(1).Text = ""
                        Combo1(1).Text = ""
                        Text2(1).Text = ""
                        Text4.Text = ""
                        Text3(2).Text = ""
                        Text5.Text = ""
                        Text6.Text = ""
                        Text7.Text = ""
                        Text8.Text = ""
                        Text9.Text = ""
                        Me.Hide
                        End If
                    Else
                    Data5.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text3(2).Text & "' or CARDNO2='" & Text3(2).Text & "' "
                    Data5.Refresh
                        If Data5.Recordset.RecordCount = 1 Then
                        MsgBox "This Card No Already Exsists", , "Error"
                        Text3(2).Text = ""
                        Else
                            If Text4.Text = Text3(2).Text Then
                            MsgBox "User Cannot Have Two Card Of Same No", , "Error"
                            Text3(2).Text = ""
                            Text4.Text = ""
                            Else
                            Data1.Recordset.Edit
                            Data1.Recordset.Update
                            MsgBox "Record Was Successfull Modified", , ""
                            Text1(1).Text = ""
                            Combo1(1).Text = ""
                            Text2(1).Text = ""
                            Text4.Text = ""
                            Text3(2).Text = ""
                            Text5.Text = ""
                            Text6.Text = ""
                            Text7.Text = ""
                            Text8.Text = ""
                            Text9.Text = ""
                            Me.Hide
                            End If
                        End If
                    End If
                Else
                Data4.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text4.Text & "' or CARDNO2='" & Text4.Text & "' "
                Data4.Refresh
                    If Data4.Recordset.RecordCount = 1 Then
                    MsgBox "This Card No Already Exsists", , "Error"
                    Text4.Text = ""
                    Else
                        If Text3(2).Text = Text7.Text Then
                            If Text4.Text = Text3(2).Text Then
                            MsgBox "User Cannot Have Two Card Of Same No", , "Error"
                            Text3(2).Text = ""
                            Text4.Text = ""
                            Else
                            Data1.Recordset.Edit
                            Data1.Recordset.Update
                            MsgBox "Record Was Successfull Modified", , ""
                            Text1(1).Text = ""
                            Combo1(1).Text = ""
                            Text2(1).Text = ""
                            Text4.Text = ""
                            Text3(2).Text = ""
                            Text5.Text = ""
                            Text6.Text = ""
                            Text7.Text = ""
                            Text8.Text = ""
                            Text9.Text = ""
                            Me.Hide
                            End If
                        Else
                        Data5.RecordSource = "SELECT * FROM CARDDETAILS WHERE CARDNO1='" & Text3(2).Text & "' or CARDNO2='" & Text3(2).Text & "' "
                        Data5.Refresh
                            If Data5.Recordset.RecordCount = 1 Then
                            MsgBox "This Card No Already Exsists", , "Error"
                            Text3(2).Text = ""
                            Else
                                If Text4.Text = Text3(2).Text Then
                                MsgBox "User Cannot Have Two Card Of Same No", , "Error"
                                Text3(2).Text = ""
                                Text4.Text = ""
                                Else
                                Data1.Recordset.Edit
                                Data1.Recordset.Update
                                MsgBox "Record Was Successfull Modified", , ""
                                Text1(1).Text = ""
                                Combo1(1).Text = ""
                                Text2(1).Text = ""
                                Text4.Text = ""
                                Text3(2).Text = ""
                                Text5.Text = ""
                                Text6.Text = ""
                                Text7.Text = ""
                                Text8.Text = ""
                                Text9.Text = ""
                                Me.Hide
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
Text1(1).Text = ""
Combo1(1).Text = ""
Text2(1).Text = ""
Text4.Text = ""
Text3(2).Text = ""
End Sub

Private Sub Command3_Click()
Text1(1).Text = ""
Combo1(1).Text = ""
Text2(1).Text = ""
Text4.Text = ""
Text3(2).Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Me.Hide
End Sub
