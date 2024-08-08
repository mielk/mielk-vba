VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EVbaVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "EVbaVariable"
'----------------------------------------------------------------------------------------------------------
Private pMethod As EVbaMethod
Private pName As String
Private pVarType As VarTypeEnum
Private pIsOptional As Boolean
Private pPassMode As ParamPassModeEnum
Private pIsArray As Boolean
Private pIsParamArray As Boolean
Private pDataType As String
Private pDefaultValue As Variant
Private pArraySize As String
'----------------------------------------------------------------------------------------------------------


'[SETTERS]
Public Function setMethod(value As EVbaMethod) As EVbaVariable
    Set setMethod = Me
    Set pMethod = value
End Function

Public Function setName(value As String) As EVbaVariable
    Set setName = Me
    pName = value
End Function

Public Function setVarType(value As VarTypeEnum) As EVbaVariable
    Set setVarType = Me
    pVarType = value
End Function

Public Function setIsOptional(value As Variant) As EVbaVariable
    Set setIsOptional = Me
    pIsOptional = F.Strings.compareStrings(value, VBA_OPTIONAL)
End Function

Public Function setPassMode(value As Variant) As EVbaVariable
    Set setPassMode = Me
    pPassMode = getParamPassModeFromName(VBA.CStr(value))
End Function

Public Function setIsArray(value As Variant) As EVbaVariable
    Set setIsArray = Me
    pIsArray = F.Strings.compareStrings(value, VBA_ARRAY_BRACKETS)
End Function

Public Function setIsParamArray(value As Variant) As EVbaVariable
    Set setIsParamArray = Me
    pIsParamArray = F.Strings.compareStrings(value, VBA_PARAM_ARRAY)
End Function

Public Function setDataType(value As Variant) As EVbaVariable
    Set setDataType = Me
    pDataType = value
End Function

Public Function setDefaultValue(value As Variant) As EVbaVariable
    Set setDefaultValue = Me
    pDefaultValue = value
End Function

Public Function setArraySize(value As String) As EVbaVariable
    Set setArraySize = Me
    pArraySize = value
End Function

Public Function loadFromDto(dto As Scripting.Dictionary) As EVbaVariable
    Set loadFromDto = Me
    Call setName(F.dictionaries.GetItem(dto, Props.name.getName))
    Call setVarType(F.dictionaries.GetItem(dto, Props_Vbe.variableType.getName))
    Call setIsOptional(F.dictionaries.GetItem(dto, Props_Vbe.IsOptional.getName))
    Call setPassMode(F.dictionaries.GetItem(dto, Props_Vbe.PassMode.getName))
    Call setIsArray(F.dictionaries.GetItem(dto, Props_Vbe.IsArray.getName))
    Call setIsParamArray(F.dictionaries.GetItem(dto, Props_Vbe.IsParamArray.getName))
    Call setDataType(F.dictionaries.GetItem(dto, Props_Vbe.DataType.getName))
    Call setDefaultValue(F.dictionaries.GetItem(dto, Props_Vbe.DefaultValue.getName))
    Call setArraySize(F.dictionaries.GetItem(dto, Props_Vbe.ArraySize.getName))
End Function




'[GETTERS]
Public Function getMethod() As EVbaMethod
    Set getMethod = pMethod
End Function

Public Function getName() As String
    getName = pName
End Function

Public Function getVarType() As VarTypeEnum
    getVarType = pVarType
End Function

Public Function getIsOptional() As Boolean
    getIsOptional = pIsOptional
End Function

Public Function getPassMode() As ParamPassModeEnum
    getPassMode = pPassMode
End Function

Public Function getIsArray() As Boolean
    getIsArray = pIsArray
End Function

Public Function getIsParamArray() As Boolean
    getIsParamArray = pIsParamArray
End Function

Public Function getDataType() As String
    getDataType = pDataType
End Function

Public Function getDefaultValue() As Variant
    getDefaultValue = pDefaultValue
End Function

Public Function getArraySize() As String
    getArraySize = pArraySize
End Function

Public Function isConst() As Boolean
    isConst = (pVarType = VarType_Const)
End Function

Public Function isNotConst() As Boolean
    isNotConst = (pVarType <> VarType_Const)
End Function

'Public Function getValue()
'
'End Function

