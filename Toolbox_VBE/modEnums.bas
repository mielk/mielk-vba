Attribute VB_Name = "modEnums"
Option Explicit

Private Const CLASS_NAME As String = "modEnums"
'----------------------------------------------------------------------------------------------------------

Public Enum ScopeTypeEnum
    ScopeType_Unknown = 0
    ScopeType_Public = 1
    ScopeType_Private = 2
    ScopeType_Friend = 3
End Enum

Public Enum VarTypeEnum
    VarType_Unknown = 0
    VarType_Dim = 1
    VarType_Const = 2
    VarType_Static = 3
    VarType_Private = 4
    VarType_Public = 5
    VarType_DimModule = 6
End Enum

Public Enum MethodTypeEnum
    MethodType_Unknown = 0
    MethodType_Sub = 1
    MethodType_Function = 2
    MethodType_PropertyLet = 3
    MethodType_PropertySet = 4
    MethodType_PropertyGet = 5
End Enum

Public Enum ParamPassModeEnum
    ParamPassMode_Unknown = 0
    ParamPassMode_ByRef = 1
    ParamPassMode_ByVal = 2
End Enum



Public Function getScopeTypeName(value As ScopeTypeEnum) As String
    Select Case value
        Case ScopeType_Public:              getScopeTypeName = VBA_PUBLIC
        Case ScopeType_Private:             getScopeTypeName = VBA_PRIVATE
        Case ScopeType_Friend:              getScopeTypeName = VBA_FRIEND
    End Select
End Function

Public Function getScopeTypeFromName(value As String) As ScopeTypeEnum
    Select Case VBA.LCase$(value)
        Case VBA.LCase$(VBA_PUBLIC):        getScopeTypeFromName = ScopeType_Public
        Case VBA.LCase$(VBA_PRIVATE):       getScopeTypeFromName = ScopeType_Private
        Case VBA.LCase$(VBA_FRIEND):        getScopeTypeFromName = ScopeType_Friend
    End Select
End Function



Public Function getVarTypeName(value As VarTypeEnum) As String
    Select Case value
        Case VarType_Dim:                   getVarTypeName = VBA_DIM
        Case VarType_Const:                 getVarTypeName = VBA_CONST
        Case VarType_Static:                getVarTypeName = VBA_STATIC
    End Select
End Function

Public Function getVarTypeFromName(value As String) As VarTypeEnum
    Select Case VBA.LCase$(value)
        Case VBA.LCase$(VBA_DIM):           getVarTypeFromName = VarType_Dim
        Case VBA.LCase$(VBA_CONST):         getVarTypeFromName = VarType_Const
        Case VBA.LCase$(VBA_STATIC):        getVarTypeFromName = VarType_Static
    End Select
End Function



Public Function getMethodTypeName(value As MethodTypeEnum) As String
    Select Case value
        Case MethodType_Sub:                getMethodTypeName = VBA_SUB
        Case MethodType_Function:           getMethodTypeName = VBA_FUNCTION
        Case MethodType_PropertySet:        getMethodTypeName = VBA_PROPERTY_SET
        Case MethodType_PropertyLet:        getMethodTypeName = VBA_PROPERTY_LET
        Case MethodType_PropertyGet:        getMethodTypeName = VBA_PROPERTY_GET
    End Select
End Function

Public Function getMethodTypeFromName(value As String) As MethodTypeEnum
    Select Case VBA.LCase$(value)
        Case VBA.LCase$(VBA_SUB):           getMethodTypeFromName = MethodType_Sub
        Case VBA.LCase$(VBA_FUNCTION):      getMethodTypeFromName = MethodType_Function
        Case VBA.LCase$(VBA_PROPERTY_SET):  getMethodTypeFromName = MethodType_PropertySet
        Case VBA.LCase$(VBA_PROPERTY_GET):  getMethodTypeFromName = MethodType_PropertyGet
        Case VBA.LCase$(VBA_PROPERTY_LET):  getMethodTypeFromName = MethodType_PropertyLet
    End Select
End Function



Public Function getParamPassModeName(value As ParamPassModeEnum) As String
    Select Case value
        Case ParamPassMode_ByRef:           getParamPassModeName = VBA_BY_REF
        Case ParamPassMode_ByVal:           getParamPassModeName = VBA_BY_VAL
    End Select
End Function

Public Function getParamPassModeFromName(value As String) As ParamPassModeEnum
    Select Case VBA.LCase$(value)
        Case VBA.LCase$(VBA_BY_REF):        getParamPassModeFromName = ParamPassMode_ByRef
        Case VBA.LCase$(VBA_BY_VAL):        getParamPassModeFromName = ParamPassMode_ByVal
    End Select
End Function
