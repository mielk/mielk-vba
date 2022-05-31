VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FConnections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FConnections"
'----------------------------------------------------------------------------------------------------------


Public Sub closeConnection(ByVal connection As Object, Optional rollbackTransaction As Boolean = False)
    If Not connection Is Nothing Then
        If TypeOf connection Is IDbConnection Then
            If rollbackTransaction Then
                Call connection.rollbackTransaction
            End If
            Call connection.closeConnection
        ElseIf isAdodbConnection(connection) Then
            If rollbackTransaction Then
                On Error Resume Next
                Call connection.RollbackTrans
                On Error GoTo 0
            End If
            Call connection.Close
        End If
    End If
    
    
ErrHandler:
    Dim errNumber As Long
    
    If rollbackTransaction And errNumber = Exceptions.dbClosingWhileInTransaction.getNumber Then
        Call VBA.Err.Clear
    Else
        'call ErrorManager ...
    End If
    
End Sub


Public Sub closeRecordset(rs As Object, Optional closeUnderlyingConnection As Boolean = True, _
                                        Optional destroyRecordset As Boolean = True)
    Dim conn As Object  'ADODB.Recordset
    '------------------------------------------------------------------------------------------------------
    
    If Not rs Is Nothing Then
        If closeUnderlyingConnection Then
            On Error Resume Next
            Set conn = rs.ActiveConnection
            Call closeConnection(conn)
            On Error GoTo 0
        End If
        
        If rs.State Then
            Call rs.Close
        End If
        
        If destroyRecordset Then Set rs = Nothing
        
    End If
    
End Sub


Private Function isAdodbConnection(value As Variant) As Boolean
    Const VBA_TYPENAME As String = "Connection"
    Const PROVIDER_TEST_INITIAL_VALUE As String = "26922CDB-2A18-4D28-A9EB-E602F80A3A63"
    '------------------------------------------------------------------------------------------------------
    Dim conn As Object
    Dim provider As String
    '------------------------------------------------------------------------------------------------------
    
    If Not VBA.IsObject(value) Then
    ElseIf value Is Nothing Then
    ElseIf F.Strings.compareStrings(VBA.TypeName(value), VBA_TYPENAME) Then
        Set conn = value
        provider = PROVIDER_TEST_INITIAL_VALUE
        
        'Try to assign ADODB.Connection provider to variable [provider].
        'If [provider] variable is still equal to PROVIDER_TEST_INITIAL_VALUE const afterwards, it
        'means that this operation failed (because it is impossible that valid ADODB.Connection has
        'this kind of provider), therefore [value] is not a valid ADODB.Connection.
        On Error Resume Next
        provider = conn.provider
        If Not dev_mode Then On Error GoTo ErrHandler
        
        isAdodbConnection = (provider <> PROVIDER_TEST_INITIAL_VALUE)
        
    End If


ErrHandler:

End Function
