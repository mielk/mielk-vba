Attribute VB_Name = "fnValidations"
'Option Explicit
'
'Private Const CLASS_NAME As String = "modValidationFunctions"
''------------------------------------------------------------------------------------------------------
'
'
''[INFRASTRUCTURE]
'Public Function getValidationFunction(name As String) As String
'    getValidationFunction = getFunctionFullName("validate_" & name, Excel.ThisWorkbook)
'End Function
'
'
'
''[FUNCTIONS]
'Public Function validate_uniqueName(values As Scripting.Dictionary) As MValidation
'    Const NOT_VALID_TAG As String = "[NameAlreadyExists]"
'    '------------------------------------------------------------------------------------------------------
'    Dim value As Variant
'    Dim existingNames As Scripting.Dictionary
'    '------------------------------------------------------------------------------------------------------
'
'    value = trimFull(stringify(getDictionaryItemByKey(values, MielkCore.VALIDATION_MAIN_VALUE)))
'    Set existingNames = getDictionaryItemByKey(values, STANDARIZATION_STANDARD_NAMES_TAG)
'    If existingNames.Exists(value) Then
'        Set validate_uniqueName = createValidation(False, NOT_VALID_TAG)
'    Else
'        Set validate_uniqueName = createValidation(True)
'    End If
'
'End Function
'
