VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------------------------------------------------------------
Private Const CLASS_NAME As String = "FVariables"
'[Constants] ----------------------------------------------------------------------------------------------
Private Const TRUE_STRING As String = "True"
Private Const FALSE_STRING As String = "False"
'----------------------------------------------------------------------------------------------------------




'----------------------------------------------------------------------------------------------------------
' Description:          Wrapper procedure to assign value to the given variable.
'                       The reason for this procedure is to solve the problem with different syntax when
'                       assigning object to a variable (use of Set keyword).
'
' Parameters:
'   * variable          Reference to the variable that given value should be assigned to.
'   * value             Value to be assigned to the given variable.
 '----------------------------------------------------------------------------------------------------------
Public Sub assign(ByRef variable As Variant, value As Variant)

    If VBA.IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If

End Sub


 '----------------------------------------------------------------------------------------------------------
' Description:          Method to swap values between two variables.
'                       Example:
'
'                                       |   variable1   |   variable2
'                           ____________________________________________
'                           before      |       1       |       2
'                           after       |       2       |       1
'
'                       If the values are not compatible, i.e. variable1 is of Object type and variable2
'                       of Long type, error will be raised (unless both of them were initially declared
'                       as Variant).
'
' Parameters:
'   * variable1, variable 2
'                       Variables which values are to be swapped.
 '----------------------------------------------------------------------------------------------------------
Public Sub swapValues(ByRef variable1 As Variant, ByRef variable2 As Variant)
    Dim tempValue As Variant
    '------------------------------------------------------------------------------------------------------

    Call f.Variables.assign(tempValue, variable1)
    Call f.Variables.assign(variable1, variable2)
    Call f.Variables.assign(variable2, tempValue)

End Sub











Public Function isDeclaredAsVariant(ByRef variable As Variant) As Boolean
    Const VT_BYREF = &H4000&
    '------------------------------------------------------------------------------------------------------
    Dim iType As Integer
    '------------------------------------------------------------------------------------------------------

    Call CopyMemory(iType, variable, 2)
    If (iType And VT_BYREF) <> VT_BYREF Then
        isDeclaredAsVariant = True
    End If

End Function

#If VBA7 Then
    Public Function getByPointer(pointer As LongPtr) As Object
#Else
    Public Function getByPointer(pointer As Long) As Object
#End If
    Set getByPointer = f.objects.getByPointer(pointer)
End Function

Public Function getObjectProperty(obj As Object, ByVal value As String) As Variant
    Dim functionName As String
    '------------------------------------------------------------------------------------------------------
    
    functionName = f.strings.substring(value, METHOD_START_TAG, METHOD_END_TAG)
    If VBA.Len(functionName) Then
        Call f.Variables.assign(getObjectProperty, VBA.CallByName(obj, functionName, VbMethod))
    Else
        getObjectProperty = value
    End If
    
End Function



'[CHECKING DATA TYPES]
Public Function isPrimitive(value As Variant) As Boolean
    If VBA.IsObject(value) Then
    ElseIf VBA.IsArray(value) Then
    ElseIf VBA.IsNull(value) Then
    Else
        isPrimitive = True
    End If
End Function

Public Function isContainer(ByRef value As Variant) As Boolean
    If VBA.IsArray(value) Then
        isContainer = (f.arrays.countDimensions(value) = 1)
    ElseIf VBA.IsObject(value) Then
        If Not value Is Nothing Then
            If TypeOf value Is VBA.Collection Then
                isContainer = True
            ElseIf TypeOf value Is Scripting.Dictionary Then
                isContainer = True
            End If
        End If
    End If
End Function


Public Function isUndefined(value As Variant, Optional includeEmptyString As Boolean = False, _
                                              Optional includeZero As Boolean = False, _
                                              Optional includeUndefinedArray As Boolean = True) As Boolean
    If VBA.IsObject(value) Then
        If value Is Nothing Then
            isUndefined = True
        End If
    ElseIf VBA.isNumeric(value) And includeZero Then
        isUndefined = (value = 0)
    ElseIf VBA.IsDate(value) And includeZero Then
        isUndefined = (value = 0)
    ElseIf VBA.IsNull(value) Then
        isUndefined = True
    ElseIf VBA.IsMissing(value) Then
        isUndefined = True
    ElseIf VBA.IsEmpty(value) Then
        isUndefined = True
    ElseIf VBA.IsArray(value) Then
        If Not f.arrays.isDefined(value) And includeUndefinedArray Then
            isUndefined = True
        End If
    ElseIf VBA.IsError(value) Then
        isUndefined = True
    ElseIf value = vbNullString And includeEmptyString Then
        isUndefined = True
    End If
End Function

Public Function isNothing(ByRef value As Variant) As Boolean
    If VBA.IsObject(value) Then
        isNothing = value Is Nothing
    End If
End Function

Public Function isAnyItemNothing(ParamArray items() As Variant) As Boolean
    Dim varItem As Variant
    '------------------------------------------------------------------------------------------------------
    For Each varItem In items
        If isNothing(varItem) Then
            isAnyItemNothing = True
            Exit For
        End If
    Next varItem
End Function


Public Function isOfNumericType(value As Variant) As Boolean
    Dim varType As VBA.VbVarType
    '------------------------------------------------------------------------------------------------------

    varType = VBA.varType(value)
    If varType = vbInteger Then
        isOfNumericType = True
    ElseIf varType = vbLong Then
        isOfNumericType = True
    ElseIf varType = vbSingle Then
        isOfNumericType = True
    ElseIf varType = vbByte Then
        isOfNumericType = True
    ElseIf varType = vbDecimal Then
        isOfNumericType = True
    ElseIf varType = vbCurrency Then
        isOfNumericType = True
    ElseIf varType = vbDouble Then
        isOfNumericType = True
    Else
        isOfNumericType = False
    End If

End Function


Public Function isString(value As Variant, Optional allowVariantWithStringAssigned As Boolean = True) As Boolean
    isString = isSpecificDataType(value, VBA.vbString, allowVariantWithStringAssigned)
End Function

Public Function isBoolean(value As Variant, Optional allowVariantWithBooleanAssigned As Boolean = True, _
                          Optional allowStringRepresentation As Boolean = False) As Boolean
    Const TRUE_STRING As String = "True"
    Const FALSE_STRING As String = "False"
    '------------------------------------------------------------------------------------------------------
    isBoolean = isSpecificDataType(value, VBA.vbBoolean, allowVariantWithBooleanAssigned)
    If Not isBoolean And allowStringRepresentation Then
        If VBA.StrComp(value, TRUE_STRING, vbTextCompare) = 0 Then
            isBoolean = True
        ElseIf VBA.StrComp(value, FALSE_STRING, vbTextCompare) = 0 Then
            isBoolean = True
        End If
    End If
End Function

Private Function isSpecificDataType(ByRef value As Variant, ByVal checkedDataType As VBA.VbVarType, _
                                                                    allowVariant As Boolean) As Boolean
    If VBA.IsObject(value) Then
        isSpecificDataType = False
    ElseIf (VBA.varType(value) = checkedDataType) Then
        If allowVariant Then
            isSpecificDataType = True
        Else
            isSpecificDataType = Not isDeclaredAsVariant(value)
        End If
    End If
End Function

Public Function isPrimitiveDataType(dataType As String) As Boolean
    Static dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If dict Is Nothing Then
        Set dict = f.dictionaries.Create(False)
        With dict
            Call .Add(TYPE_STRING, vbNullString)
            Call .Add(TYPE_BYTE, vbNullString)
            Call .Add(TYPE_INTEGER, vbNullString)
            Call .Add(TYPE_LONG, vbNullString)
            Call .Add(TYPE_DOUBLE, vbNullString)
            Call .Add(TYPE_DATE, vbNullString)
            Call .Add(TYPE_BOOLEAN, vbNullString)
            Call .Add(TYPE_CURRENCY, vbNullString)
        End With
    End If
    
    isPrimitiveDataType = dict.exists(dataType)
    
End Function








'[COMPARING OBJECTS]
 '----------------------------------------------------------------------------------------------------------
' Description:          Function to check if two given values [a] and [b] are equal.
' Returned value:       True - in the following cases:
'                           * both values are initialized arrays and they have:
'                                   - the same number of dimension,
'                                   - the same size for each dimension,
'                                   - the same indexing base for each dimension (unless parameter
'                                     [ignoreIndexingBase] is set to True),
'                                   - equal value in each cell,
'                                   - value of the same data type in each cell (unless parameter
'                                     [strictDataType] is set to False),
'                           * both values are uninitalized dynamic arrays and have the same data type
'                                   (unless [strictDataType] parameter is set to False),
'                           * both values are Null,
'                           * both values are VBA.Missing,
'                           * both values are VBA.Empty,
'                           * both values are Nothing,
'                           * both values are of different undefined type (Null/Missing/Empty/Nothing) and
'                                   [strictDataType] is set to False,
'                           * one value is of undefined type (Null/Missing/Empty/Nothing) and second one is
'                                   empty string / False / 0 and [strictDataType] is set to False,
'                           * both values are of Dictionary type and have the same key-item pairs
'                                   Note that:
'                                       - if parameter [strictDataType] is set to True, compared keys and
'                                         items must have the same data type,
'                                       - if parameter [isCaseSensitive] is set to True, string keys and
'                                         items must have the same letter casing when compared,
'                           * both values are of Collection type and contain the same items
'                                   Note that:
'                                       - items don't need to be stored in the same order, unless parameter
'                                         [orderSensitiveContainers] is set to True,
'                                       - if parameter [strictDataType] is set to True, compared items must
'                                         have the same data type,
'                                       - if parameter [isCaseSensitive] is set to True, string items must
'                                         have the same letter casing when compared,
'                           * both values are references to the same non-container object,
'                           * value [a] is an object of a class that have its own [equals] method and this
'                                   method returns True when value [b] is passed,
'                           * both values are of primitive type, they have the same value and are of the
'                                   same data type (unless parameter [strictDataType] is set to False)
'                                   Note that:
'                                       - when comparing two strings, they need to be in the same letter
'                                         casing (unless parameter [isCaseSensitive] is set to False),
'                                       - if parameter [trimWhenComparingStrings] is set to True, blank
'                                         spaces at the beginning and at the end of compared strings are
'                                         ignored.
'
'                       False - in any other case.
'
'
' Parameters:
'   * a, b              Compared values.
'   * strictDataType    (Optional) If this value is set to True, both compared value must have exactly the
'                       the same data type. Otherwise they can be of different data type and the function
'                       still can return false if they have the same value, i.e. "2" and 2.
'                       It applies for each item when containers are compared.
'   * isCaseSensitive   (Optional) If this value is set to True and string are compared they need to be in
'                       the same letter casing to be considered as equal. Otherwise comparing is
'                       case-insensitive (i.e. "abc" is equal "ABC").
'                       It applies for each item when containers are compared.
'   * trimWhenComparingStrings
'                       (Optional) If this value is set to False and string are compared all blank spaces
'                       at the beginning and the end of those strings are ignored (i.e. "  abc  " is equal
'                       to "abc"). Otherwise blank spaces are taken into account when comparing strings.
'   * ignoreIndexingBase
'                       (Optional) If this parameter is set to True and arrays are compared then it doesn't
'                       matter what is base index for each dimension of that array. Otherwise, both arrays
'                       need to have the same base index for each dimension.
'   * orderSensitiveForContainers
'                       (Optional) If this value is set to True and containers are compared, all items in
'                       those containers need to be placed in exact the same order. Otherwise, order of
'                       items is ignored (i.e. Collection [1, 2, 3] would be considered as equal to
'                       Collection [3, 1 ,2]).
 '----------------------------------------------------------------------------------------------------------
Public Function equals(ByVal a As Variant, ByVal b As Variant, _
                                Optional strictDataType As Boolean = False, _
                                Optional isCaseSensitive As Boolean = False, _
                                Optional trimWhenComparingStrings As Boolean = True, _
                                Optional ignoreIndexingBase As Boolean = True, _
                                Optional orderSensitiveForContainers As Boolean = True) As Boolean

    If VBA.IsArray(a) And VBA.IsArray(b) Then
        equals = f.arrays.compare(a, b, ignoreIndexingBase, strictDataType)
    ElseIf VBA.IsArray(a) Or VBA.IsArray(b) Then
        equals = False
    ElseIf f.Variables.isUndefined(a) And isUndefined(b) Then
        equals = equals_checkForBothUndefined(a, b, strictDataType)
    ElseIf f.Variables.isUndefined(a) Then
        equals = equals_checkForOneUndefined(a, b, strictDataType)
    ElseIf f.Variables.isUndefined(b) Then
        equals = equals_checkForOneUndefined(b, a, strictDataType)
    ElseIf VBA.IsObject(a) And VBA.IsObject(b) Then
        equals = equals_checkForObjects(a, b, strictDataType, isCaseSensitive, orderSensitiveForContainers)
    ElseIf VBA.IsObject(a) Or VBA.IsObject(b) Then
        equals = False
    ElseIf f.Dates.isDateString(a) <> f.Dates.isDateString(b) Then
        equals = False
    Else
        equals = equals_checkForPrimitives(a, b, strictDataType, isCaseSensitive, trimWhenComparingStrings)
    End If

End Function

Private Function equals_checkForBothUndefined(ByVal a As Variant, ByVal b As Variant, _
                                                                      strictDataType As Boolean) As Boolean
    If strictDataType Then
        If VBA.IsMissing(a) And VBA.IsMissing(b) Then
            equals_checkForBothUndefined = True
        ElseIf VBA.IsNull(a) And VBA.IsNull(b) Then
            equals_checkForBothUndefined = True
        ElseIf VBA.IsEmpty(a) And VBA.IsEmpty(b) Then
            equals_checkForBothUndefined = True
        ElseIf isNothing(a) And isNothing(b) Then
            equals_checkForBothUndefined = True
        End If
    Else
        equals_checkForBothUndefined = True
    End If
End Function

Private Function equals_checkForOneUndefined(ByVal undefined As Variant, ByVal defined As Variant, _
                                                                      strictDataType As Boolean) As Boolean
    If VBA.IsEmpty(undefined) Then
        If Not strictDataType Then
            If isPrimitive(defined) Then
                If defined = vbNullString Then
                    equals_checkForOneUndefined = True
                ElseIf defined = False Then
                    equals_checkForOneUndefined = True
                ElseIf defined = 0 Then
                    equals_checkForOneUndefined = True
                End If
            End If
        End If
    End If
End Function

Private Function equals_checkForObjects(ByVal a As Variant, ByVal b As Variant, _
                strictDataType As Boolean, caseSensitive As Boolean, orderSensitive As Boolean) As Boolean
    Const METHOD_NAME As String = "equals_checkForObjects"
    '------------------------------------------------------------------------------------------------------
    Const EQUALS_METHOD_NAME As String = "equals"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If a Is b Then
        equals_checkForObjects = True
    ElseIf f.Collections.isCollection(a) And f.Collections.isCollection(b) Then
        equals_checkForObjects = f.Collections.compare(f.cast.toCollection(a), f.cast.toCollection(b), _
                                                            strictDataType, orderSensitive, caseSensitive)
    ElseIf f.dictionaries.isDictionary(a) And f.dictionaries.isDictionary(b) Then
        equals_checkForObjects = f.dictionaries.compare(f.cast.toDictionary(a), f.cast.toDictionary(b), _
                                            strictDataType:=strictDataType, isCaseSensitive:=caseSensitive)
    Else
        Call VBA.err.clear
        On Error Resume Next
        equals_checkForObjects = a.equals(b)    'VBA.CallByName(a, EQUALS_METHOD_NAME, VbMethod, b)
        If VBA.err.number = Exceptions.MethodNotSupported.getNumber Then
            Call VBA.err.clear
        End If
        If Not DEV_MODE Then On Error GoTo ErrHandler
    End If

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = VBA.err.number: errDescription = VBA.err.description
    If errNumber <> Exceptions.MethodNotSupported.getNumber Then
        Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
                        "a; b; strictDataType; caseSensitive; orderSensitive", _
                         a, b, strictDataType, caseSensitive, orderSensitive)
    Else
        errNumber = 0
    End If
    GoTo ExitPoint

End Function

Private Function equals_checkForArrays(ByVal a As Variant, ByVal b As Variant, _
                                    ignoreIndexingBase As Boolean, strictDataType As Boolean) As Boolean
    Const METHOD_NAME As String = "equals_checkForArrays"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    equals_checkForArrays = f.arrays.compare(a, b, ignoreIndexingBase, strictDataType)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

 '----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = VBA.err.number: errDescription = VBA.err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
                        "a; b; ignoreIndexingBase; strictDataType", _
                         a, b, ignoreIndexingBase, strictDataType)
    GoTo ExitPoint

End Function

Private Function equals_checkForPrimitives(ByVal baseValue As Variant, ByVal comparedValue As Variant, _
                            strictDataType As Boolean, isCaseSensitive As Boolean, _
                            trimStrings As Boolean) As Boolean
    Const METHOD_NAME As String = "equals_checkForPrimitives"
    '------------------------------------------------------------------------------------------------------

    If haveTheSameDataType(baseValue, comparedValue) Then
        If VBA.varType(baseValue) = vbString Then
            equals_checkForPrimitives = f.strings.compareStrings(f.strings.stringify(baseValue), _
                            f.strings.stringify(comparedValue), isCaseSensitive, trimStrings)
        Else
            equals_checkForPrimitives = (baseValue = comparedValue)
        End If
    ElseIf strictDataType Then
        equals_checkForPrimitives = False
    ElseIf isBoolean(baseValue) Then
        equals_checkForPrimitives = equals_checkForBoolean(baseValue, comparedValue, trimStrings)
    ElseIf isBoolean(comparedValue) Then
        equals_checkForPrimitives = equals_checkForBoolean(comparedValue, baseValue, trimStrings)
    Else
        equals_checkForPrimitives = f.strings.compareStrings(f.strings.stringify(baseValue), _
                                        f.strings.stringify(comparedValue), isCaseSensitive, trimStrings)
    End If

End Function

Private Function equals_checkForBoolean(booleanValue As Variant, nonBooleanValue As Variant, _
                                                                                trim As Boolean) As Boolean
    Const METHOD_NAME As String = "equals_checkForBoolean"
    '------------------------------------------------------------------------------------------------------

    If isString(nonBooleanValue) Then
        equals_checkForBoolean = f.strings.compareStrings(VBA.IIf(booleanValue, TRUE_STRING, FALSE_STRING), _
                                                                            VBA.CStr(nonBooleanValue), False, trim)
    ElseIf isNumeric(nonBooleanValue) Then
        equals_checkForBoolean = (booleanValue = VBA.CBool(nonBooleanValue))
    End If

End Function


Private Function haveTheSameDataType(ByRef a As Variant, ByRef b As Variant) As Boolean
    haveTheSameDataType = (VBA.varType(a) = VBA.varType(b))
End Function







Public Function coalesce(ParamArray values() As Variant) As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    For Each value In values
        If Not f.Variables.isUndefined(value, True, True, True) Then
            Call f.Variables.assign(coalesce, value)
            Exit For
        End If
    Next value
End Function



'
'
'
'
'
'
'
''[MissingUnitTests]
' '----------------------------------------------------------------------------------------------------------
'' Description:          Function compares two values [baseValue] and [comparedValue] and returns -1, 0 or 1
''                       depending on which value is greater.
''                       The rules used to determine which value is greater are described further in
''                       this comment.
''
''                       This function is of great use especially when sorting data in containers.
''
'' Returned value:       If value [baseValue] is less than value [comparedValue] function returns -1.
''                       If both values are equal function returns 0.
''                       If value [baseValue] is greater than value [comparedValue] function returns 1.
''
''                       Rules used to determine which value is greater:
''                           *
''
'' Parameters:
''   * baseValue, comparedValue
''                       Values that are being compared.
''   * strictType        (Optional) If this value is set to False, data type is ignored when comparing
''                       values, i.e. 1 is greater than "0". Otherwise, data type is taken into
''                       consideration and "0" is greater than 1 (because strings are considered to be
''                       greater than numbers).
''   * isCaseSensitive   (Optional) If this value is set to False, letter casing is ignored when comparing
''                       string values, i.e. "B" is greater than "a". Otherwise, letter casing is taken
''                       into account and "a" is greater than "B" (because lower case letter are considered
''                       to be greater than upper case letter, since they have higher ASCII codes).
' '----------------------------------------------------------------------------------------------------------
'Public Function compare(baseValue As Variant, comparedValue As Variant, _
'                        Optional strictType As Boolean = True, _
'                        Optional isCaseSensitive As Boolean = True) As Integer
'    Const METHOD_NAME As String = "compare"
'    '------------------------------------------------------------------------------------------------------
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'    If F.Variables.isUndefined(baseValue, False, False, False) Then
'        compare = VBA.IIf(F.Variables.isUndefined(comparedValue, False, False, False), 0, 1)
'    ElseIf F.Variables.isUndefined(comparedValue, False, False, False) Then
'        compare = -1
'    ElseIf VBA.IsObject(baseValue) Then
'        compare = VBA.IIf(VBA.IsObject(comparedValue), 0, 1)
'    ElseIf VBA.IsObject(comparedValue) Then
'        compare = -1
'    ElseIf VBA.IsArray(baseValue) Then
'        compare = VBA.IIf(VBA.IsArray(comparedValue), 0, 1)
'    ElseIf VBA.IsArray(comparedValue) Then
'        compare = -1
'    Else
'        compare = comparePrimitives(baseValue, comparedValue, strictType, isCaseSensitive)
'    End If
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
' '----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("baseValue", baseValue), _
'                    KeyValue("comparedValue", comparedValue), KeyValue("strictType", strictType), _
'                    KeyValue("isCaseSensitive", isCaseSensitive))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function
'
'Private Function comparePrimitives(baseValue As Variant, comparedValue As Variant, _
'                        Optional strictType As Boolean = True, _
'                        Optional isCaseSensitive As Boolean = True) As Integer
'    Const METHOD_NAME As String = "comparePrimitives"
'    '------------------------------------------------------------------------------------------------------
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'
'    If strictType Then
'        comparePrimitives = comparePrimitives_strictTypes(baseValue, comparedValue, isCaseSensitive)
'    Else
'        comparePrimitives = comparePrimitives_notStrictTypes(baseValue, comparedValue, isCaseSensitive)
'    End If
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
' '----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("baseValue", baseValue), _
'                    KeyValue("comparedValue", comparedValue), KeyValue("strictType", strictType), _
'                    KeyValue("isCaseSensitive", isCaseSensitive))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function
'
'Private Function comparePrimitives_strictTypes(baseValue As Variant, comparedValue As Variant, _
'                        Optional isCaseSensitive As Boolean = True) As Integer
'    Const METHOD_NAME As String = "comparePrimitives_strictTypes"
'    '------------------------------------------------------------------------------------------------------
'    Dim baseType As VBA.VbVarType
'    Dim comparedType As VBA.VbVarType
'    '------------------------------------------------------------------------------------------------------
'
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'
'    baseType = VBA.varType(baseValue)
'    comparedType = VBA.varType(comparedValue)
'
'    If baseType = vbString Then
'        If comparedType = vbString Then
'            comparePrimitives_strictTypes = VBA.StrComp(baseValue, comparedValue, _
'                                                VBA.IIf(isCaseSensitive, vbBinaryCompare, vbTextCompare))
'        Else
'            comparePrimitives_strictTypes = 1
'        End If
'    ElseIf baseType = vbBoolean Then
'        If comparedType = vbString Then
'            comparePrimitives_strictTypes = -1
'        ElseIf comparedType = vbBoolean Then
'            If baseValue = comparedValue Then
'                comparePrimitives_strictTypes = 0
'            Else
'                comparePrimitives_strictTypes = VBA.IIf(baseValue, 1, -1)
'            End If
'        Else
'            comparePrimitives_strictTypes = 1
'        End If
'    ElseIf baseValue > comparedValue Then
'        comparePrimitives_strictTypes = 1
'    ElseIf baseValue = comparedValue Then
'        comparePrimitives_strictTypes = 0
'    Else
'        comparePrimitives_strictTypes = -1
'    End If
'
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
' '----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("baseValue", baseValue), _
'                    KeyValue("comparedValue", comparedValue), KeyValue("isCaseSensitive", isCaseSensitive))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function
'
'Private Function comparePrimitives_notStrictTypes(baseValue As Variant, comparedValue As Variant, _
'                        Optional isCaseSensitive As Boolean = True) As Integer
'    Const METHOD_NAME As String = "comparePrimitives_notStrictTypes"
'    '------------------------------------------------------------------------------------------------------
'    Dim realBaseValue As Variant
'    Dim realComparedValue As Variant
'    '------------------------------------------------------------------------------------------------------
'
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'
'    realBaseValue = comparePrimitives_convertToNonStrictDataTypeValue(baseValue)
'    realComparedValue = comparePrimitives_convertToNonStrictDataTypeValue(comparedValue)
'    comparePrimitives_notStrictTypes = comparePrimitives_strictTypes(realBaseValue, realComparedValue)
'
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
' '----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("baseValue", baseValue), _
'                    KeyValue("comparedValue", comparedValue), KeyValue("isCaseSensitive", isCaseSensitive))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function
'
'Private Function comparePrimitives_convertToNonStrictDataTypeValue(value As Variant) As Variant
'    Const METHOD_NAME As String = "comparePrimitives_convertToNonStrictDataTypeValue"
'    '------------------------------------------------------------------------------------------------------
'
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'
'    If VBA.varType(value) = vbBoolean Then
'        comparePrimitives_convertToNonStrictDataTypeValue = VBA.CBool(value)
'    ElseIf VBA.varType(value) = vbDate Then
'        comparePrimitives_convertToNonStrictDataTypeValue = VBA.CDate(value)
'    ElseIf VBA.varType(value) = vbString Then
'        If F.Numerics.isNumeric(value) Then
'            comparePrimitives_convertToNonStrictDataTypeValue = F.Numerics.toNumeric(value)
'        ElseIf F.Dates.isDateString(value) Then
'            comparePrimitives_convertToNonStrictDataTypeValue = F.Dates.fromString(value)
'        ElseIf isBoolean(value, True, True) Then
'            comparePrimitives_convertToNonStrictDataTypeValue = VBA.CBool(value)
'        Else
'            comparePrimitives_convertToNonStrictDataTypeValue = value
'        End If
'    Else
'        comparePrimitives_convertToNonStrictDataTypeValue = value
'    End If
'
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
' '----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("value", value))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function



Public Function isEqualToAny(baseItem As Variant, ParamArray compared() As Variant) As Boolean
    Const METHOD_NAME As String = "isEqualToAny"
    '------------------------------------------------------------------------------------------------------
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    For Each value In compared
        If equals(baseItem, value) Then
            isEqualToAny = True
            Exit For
        End If
    Next value
    
End Function
