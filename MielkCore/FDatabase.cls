VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FDatabase"

Public Sub closeConnection(connection As Object, Optional rollbackTransaction As Boolean = False)
    If Not connection Is Nothing Then
        If rollbackTransaction Then
            On Error Resume Next
            Call connection.RollbackTrans
            On Error GoTo 0
        End If
        If connection.state Then
            Call connection.Close
        End If
    End If
End Sub

Public Sub closeRecordset(rs As Object, Optional closeUnderlyingConnection As Boolean = True, _
                                        Optional destroyRecordset As Boolean = True)
    Dim conn As Object 'ADODB.Connection
    '------------------------------------------------------------------------------------------------------
    
    If Not rs Is Nothing Then
        If closeUnderlyingConnection Then
            On Error Resume Next
            Set conn = rs.ActiveConnection
            Call closeConnection(conn)
            On Error GoTo 0
        End If
        
        If rs.state Then
            Call rs.Close
        End If
        If destroyRecordset Then Set rs = Nothing
    End If
    
End Sub


Public Function createSqlInsertCommand(tableName As String, Items As Scripting.Dictionary) As String
    Const SQL_PATTERN As String = "INSERT INTO {0}({1}) SELECT {2}"
    '------------------------------------------------------------------------------------------------------
    Dim Key As Variant
    Dim fields As VBA.Collection
    Dim values As VBA.Collection
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If f.dictionaries.isNonEmptyDictionary(Items) Then
        Set fields = New VBA.Collection
        Set values = New VBA.Collection
        For Each Key In Items
            Call fields.add(Key)
            Call values.add(f.strings.toDbValue(Items.Item(Key)))
        Next Key
        createSqlInsertCommand = f.strings.format(SQL_PATTERN, tableName, f.Collections.toString(fields), f.Collections.toString(values))
    End If
    
End Function
