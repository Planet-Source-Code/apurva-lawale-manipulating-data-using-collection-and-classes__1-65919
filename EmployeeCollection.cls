VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'We have defined employees as a collection
Private m_Employees As Collection

'Add a new employee to the collection
Public Sub Add(ByVal emp As Employee, ByVal key As Variant)
        m_Employees.Add emp, key
End Sub

'Clear the entire collection
Public Sub Clear()
    Set m_Employees = New Collection
End Sub

'Count the number of users in a collection and return the count
Public Function Count() As Long
    Count = m_Employees.Count
End Function

'Return an employee record from a collection
Public Function Item(ByVal Index As Variant) As Employee
Attribute Item.VB_UserMemId = 0
    Set Item = m_Employees(Index)
End Function

'Delete a record from the collection
Public Sub Remove(ByVal Index As Variant)
    m_Employees.Remove Index
End Sub

'Enumerate the records in a collection
'This function is for the FOR EACH used in the
'ListEmployees function. If this function was not used
'then you can't enumerate the records in a collection
'as shown in the ListEmployees function in FORM1 using FOR EACH
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_Employees.[_NewEnum]
End Function

'Initialise the Collection
Private Sub Class_Initialize()
    Set m_Employees = New Collection
End Sub
