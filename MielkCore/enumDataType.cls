VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "enumDataType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "enumDataType"
'----------------------------------------------------------------------------------------------------------
Private pName As String
Private pReadDbTable As String
Private pWriteDbTable As String
'----------------------------------------------------------------------------------------------------------


'[SETTERS]
Public Function setName(value As String) As enumDataType
    Set setName = Me
    pName = value
End Function

Public Function setDbTable(value As String) As enumDataType
    Set setDbTable = Me
    pReadDbTable = value
    pWriteDbTable = value
End Function

Public Function setReadDbTable(value As String) As enumDataType
    Set setReadDbTable = Me
    pReadDbTable = value
End Function

Public Function setWriteDbTable(value As String) As enumDataType
    Set setWriteDbTable = Me
    pWriteDbTable = value
End Function



'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getReadDbTable() As String
Attribute getReadDbTable.VB_UserMemId = 0
    getReadDbTable = pReadDbTable
End Function

Public Function getWriteDbTable() As String
    getWriteDbTable = pWriteDbTable
End Function

Public Function toString() As String
    toString = pName
End Function
