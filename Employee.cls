VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Think of a collection as a database file.
'If you did that then the variables below are the fields of a table
'If you think of the variables below as fields then when used together
'they form a record.
'Simple to imagine now huh.
'Now how do we deal with the record, for that we have EmployeeCollection
Public LastName As String
Public FirstName As String
Public EmployeeId As Long
