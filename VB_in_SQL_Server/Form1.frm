VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "VB with MS SQL Server Database"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Reload"
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5530
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Find"
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add new"
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtPosition 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call clear_text
Call enable_Text
txtName.SetFocus
Call CommandActivate(True, True, False, False, True, True)
End Sub

Private Sub Command2_Click()
nam = Trim(txtName.Text)
pos = Trim(txtPosition.Text)

SQLQuery = "Insert Into employee values('" & nam & "','" & pos & "')"
SQLcon.Execute (SQLQuery)
nam = ""
pos = ""
Call clear_text
Call disable_Text
Call CommandActivate(True, False, False, False, True, False)

SQLQuery = "Select*from employee ORDER BY EMPName ASC"
Set Grid.DataSource = SQLcon.Execute(SQLQuery)

MsgBox "New entry has been saved.", vbInformation, "Program SQL Server"

End Sub

Private Sub Command3_Click()
nam = Trim(txtName.Text)
pos = Trim(txtPosition.Text)

SQLQuery = "Update employee set EMPName='" & nam & "', EMPPosition='" & pos & "' where EMPID='" & Idx & "'"
SQLcon.Execute (SQLQuery)
nam = ""
pos = ""
Idx = ""

Call clear_text
Call enable_Text
Call CommandActivate(True, False, False, False, True, False)

SQLQuery = "Select*from employee ORDER BY EMPName ASC"
Set Grid.DataSource = SQLcon.Execute(SQLQuery)

MsgBox "Entry has been updated successfully."
End Sub

Private Sub Command4_Click()
SQLQuery = "Delete from employee where EMPID='" & Idx & "'"
SQLcon.Execute (SQLQuery)
Idx = ""
Call clear_text
Call disable_Text
Call CommandActivate(True, False, False, False, True, False)

SQLQuery = "Select*from employee ORDER BY EMPName ASC"
Set Grid.DataSource = SQLcon.Execute(SQLQuery)

MsgBox "Entry has been deleted."
End Sub

Private Sub Command5_Click()
Idx = InputBox("Please Enter EmployeeID to delete.")
SQLQuery = "Select* from employee where EMPID='" & Idx & "'"
Set Grid.DataSource = SQLcon.Execute(SQLQuery)
End Sub

Private Sub Command6_Click()
Call clear_text
Call disable_Text
Call CommandActivate(True, False, False, False, True, False)
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Command8_Click()
SQLQuery = "Select*from employee ORDER BY EMPName ASC"
Set Grid.DataSource = SQLcon.Execute(SQLQuery)
End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Please configure these server parameter to suite your SQL Server environment'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dbUser = "hackeem"              'This is a MS SQL Server User Login Name

dbPassword = "kilburn"          'Login Password

dbName = "programDB"            'MS SQL Server Database Name

dbServer = "192.168.1.108"      'This is the Host computer on a network
                                ' You may change it into localhost if you are running on a server.
                                ' or if you are on a network use the computer name or IP address where MS SQL Server resides.

SQLcon.Open "Provider=SQLOLEDB.1; User ID=" & dbUser & ";Password=" & dbPassword & ";Initial Catalog=" & dbName & "; Data Source=" & dbServer

SQLQuery = "Select*from employee ORDER BY EMPName ASC"
Set Grid.DataSource = SQLcon.Execute(SQLQuery)

Call disable_Text
Call CommandActivate(True, False, False, False, True, False)

End Sub


''''''''FUNCTIONS''''''''''''
Private Function clear_text()
txtName.Text = ""
txtPosition.Text = ""
End Function
Private Function enable_Text()
txtName.Enabled = True
txtPosition.Enabled = True
End Function
Private Function disable_Text()
txtName.Enabled = False
txtPosition.Enabled = False
End Function


Public Function CommandActivate(cmd1, cmd2, cmd3, cmd4, cmd5, cmd6)
    Command1.Enabled = cmd1
    Command2.Enabled = cmd2
    Command3.Enabled = cmd3
    Command4.Enabled = cmd4
    Command5.Enabled = cmd5
    Command6.Enabled = cmd6
End Function


Private Sub Grid_DblClick()
Call clear_text
Call enable_Text
txtName.SetFocus
Call CommandActivate(False, False, True, True, True, True)

x = Grid.Row
With Grid
    Idx = .TextMatrix(x, 1)
    txtName.Text = .TextMatrix(x, 2)
    txtPosition.Text = .TextMatrix(x, 3)
End With
Command3.Caption = "Update"
End Sub
