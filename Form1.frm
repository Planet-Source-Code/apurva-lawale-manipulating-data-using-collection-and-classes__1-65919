VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employees - Database using Collections"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtId 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton cmdRemoveByIndex 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0000
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You can use this code in whatever way you want.
'To handle a lot of data newbies normally use arrays so I hope this example
'will help them understand Collections + Classes and appreciate their worth.
'You could use the technique showed in this example to manipulate data in memory
'then you could save it in an XML file or database.
'You could load database table info and manipulate it / update it.
'You could do a lot with much less code without breaking a sweat.

'The code in Employee and EmployeeCollection classes are commented.
'Understand those two classes. They are pretty small.
'The comments should help you understand them well.
'I recommend you look at the Employee Class first and understand
'Then look at the EmployeeCollection Class
'Then read the code in this form.

'Vote if you like this small elegant solution. Voting much appreciated.
'Apurva Lawale - http://www.apurvalawale.com/

Option Explicit

Private m_Employees As EmployeeCollection

Private Sub ListEmployees()
Dim emp As Employee

List1.Clear
For Each emp In m_Employees
    List1.AddItem emp.LastName & ", " & emp.FirstName
Next emp

End Sub

Private Sub LoadData()
Dim emp As Employee

    Set m_Employees = New EmployeeCollection

    Set emp = New Employee
    With emp
        .FirstName = "Andrew"
        .LastName = "Poulos"
        .EmployeeId = 1
    End With
    m_Employees.Add emp, emp.LastName & "," & emp.FirstName

    Set emp = New Employee
    With emp
        .FirstName = "Apurva"
        .LastName = "Lawale"
        .EmployeeId = 2
    End With
    m_Employees.Add emp, emp.LastName & "," & emp.FirstName

    Set emp = New Employee
    With emp
        .FirstName = "Mike"
        .LastName = "Baker"
        .EmployeeId = 3
    End With
    m_Employees.Add emp, emp.LastName & "," & emp.FirstName

End Sub
' Add a new employee using LastName,FirstName
' as the key.
Private Sub cmdAdd_Click()
Dim emp As New Employee

    emp.FirstName = txtFirstName.Text
    emp.LastName = txtLastName.Text
    emp.EmployeeId = txtId.Text
    On Error GoTo AddError
    m_Employees.Add emp, emp.LastName & "," & emp.FirstName
    On Error GoTo 0

    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtId.Text = ""
    ListEmployees
    Exit Sub

AddError:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdClear_Click()
    m_Employees.Clear
    ListEmployees
End Sub

Private Sub cmdNew_Click()
txtFirstName.Text = ""
txtLastName.Text = ""
txtId.Text = m_Employees.Count + 1
End Sub

Private Sub cmdRemoveByIndex_Click()
    On Error GoTo RemoveByIndexError
    m_Employees.Remove (List1.ListIndex + 1)
    On Error GoTo 0

    ListEmployees
    Exit Sub

RemoveByIndexError:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
    LoadData
    ListEmployees
End Sub

Private Sub List1_Click()
Dim emp As Employee

    On Error GoTo GetByIndexError
    Set emp = m_Employees(List1.ListIndex + 1)
    On Error GoTo 0

    txtFirstName.Text = emp.FirstName
    txtLastName.Text = emp.LastName
    txtId.Text = emp.EmployeeId
    Exit Sub

GetByIndexError:
    MsgBox Err.Description
    Exit Sub
End Sub
