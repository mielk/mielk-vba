VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FStrings"
'----------------------------------------------------------------------------------------------------------





'[STATE CHECKING]

'----------------------------------------------------------------------------------------------------------
' Description:          Function checks if the given value is not empty string.
' Returned value:       True - in the following cases:
'                           1) if the given value is of String type and has at least one character,
'                           2) if the given value is of any primitive value (other than  vbNullString,
'                              Empty, Null, Missing) and parameter [strictType] is set to False.
'                       False - in the following cases:
'                           1) the given value is object or array,
'                           2) the given value is undefined (Empty / Null / Missing / Nothing),
'                           3) the given value is of other type than String and parameter [strictType]
'                              is set to True,
'                           4) the given value is of string type but its length is 0.
'
' Parameters:
'   * value             Value to be checked.
'   * strictType        (Optional) If this value is set to True the given value must be of String type.
'                       Otherwise the value is casted to String before being checked (unless it is
'                       impossible, i.e. for arrays or objects).
'----------------------------------------------------------------------------------------------------------
Public Function isNotEmptyString(value As Variant, Optional strictType As Boolean = False) As Boolean
    Const METHOD_NAME As String = "isNotEmptyString"
    '------------------------------------------------------------------------------------------------------
    Dim valueLength As Long
    Dim varValueType As VbVarType
    '------------------------------------------------------------------------------------------------------

    On Error Resume Next
    valueLength = VBA.Len(value)
    
    'If Not DEV_MODE Then On Error GoTo ErrHandler

    If valueLength > 0 Then
        varValueType = VBA.varType(value)
        If varValueType = VBA.vbBoolean Then
            isNotEmptyString = False
        ElseIf Not strictType Then
            isNotEmptyString = True
        ElseIf varValueType = VBA.vbString Then
            isNotEmptyString = True
        End If
    End If

End Function



Public Function startsWith(text As String, prefix As String, Optional isCaseSensitive As Boolean = False) As Boolean
    Dim length As Long
    '------------------------------------------------------------------------------------------------------
    length = VBA.Len(prefix)
    startsWith = compareStrings(VBA.left$(text, length), prefix, isCaseSensitive)
End Function

Public Function endsWith(text As String, substr As String, Optional isCaseSensitive As Boolean = False) As Boolean
    Dim length As Long
    '------------------------------------------------------------------------------------------------------
    length = VBA.Len(substr)
    endsWith = compareStrings(VBA.right$(text, length), substr, isCaseSensitive)
End Function






'[COMPARING]

'----------------------------------------------------------------------------------------------------------
' Description:          Function compares two given strings.
'
' Returned value:       True - if the given strings are equal.
'                       False - otherwise.
'
' Parameters:
'   * baseString, comparedString
'                       Strings to be compared.
'   * isCaseSensitive   (Optional) If this parameter is set to True, both string have to be in the same
'                       letter casing.
'   * trimmed           (Optional) If this parameter is set to True, redundant blank characters are ignored
'                       when comparing.
'
' Exceptions:
'   IllegalDataType     Thrown if any of the given values to be compared is not of a primitive type.
'----------------------------------------------------------------------------------------------------------
Public Function compareStrings(baseString As Variant, comparedString As Variant, _
                                Optional isCaseSensitive As Boolean = False, _
                                Optional trimmed As Boolean = True) As Boolean
    Const METHOD_NAME As String = "compareStrings"
    '------------------------------------------------------------------------------------------------------
    Dim compareMethod As VBA.VbCompareMethod
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If trimmed Then
        compareStrings = (VBA.StrComp(VBA.trim(baseString), VBA.trim(comparedString), compareMethod) = 0)
    Else
        compareStrings = (VBA.StrComp(baseString, comparedString, compareMethod) = 0)
    End If
    
'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
                        "baseString; comparedString; isCaseSensitive; trimmed", _
                         baseString, comparedString, isCaseSensitive, trimmed)
    GoTo ExitPoint

End Function







'[CONVERTING]
Public Function format(pattern As String, ParamArray replacements() As Variant) As String
    Const METHOD_NAME As String = "format"
    '------------------------------------------------------------------------------------------------------
    Const SLASH_TEMP_REPLACEMENT As String = "$%^#$%^#$%"
    '------------------------------------------------------------------------------------------------------
    Dim arr As Variant
    Dim Index As Long
    Dim tag As String
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    'New line symbols
    format = VBA.replace(pattern, "\\", SLASH_TEMP_REPLACEMENT)
    format = VBA.replace(format, NEW_LINE_SYMBOL, VBA.vbCrLf)
    format = VBA.replace(format, SLASH_TEMP_REPLACEMENT, "\")

    'Replacements
    arr = f.arrays.flatArray(f.arrays.fromParamArray(replacements))
    
    For i = LBound(arr) To UBound(arr)
        Index = i - LBound(arr)
        tag = "{" & Index & "}"
        format = VBA.replace(format, tag, arr(i))
    Next i

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
                        "pattern; replacements; arr; index; tag", _
                         pattern, replacements, arr, Index, tag)
    GoTo ExitPoint

End Function



Public Function toDbValue(value As Variant, Optional throwErrorIfIllegalDataType As Boolean = True) As String
    Const METHOD_NAME As String = "toDbValue"
    '------------------------------------------------------------------------------------------------------

    If f.Variables.isUndefined(value) Then
        toDbValue = "NULL"
    ElseIf VBA.IsObject(value) Then
        If throwErrorIfIllegalDataType Then Call Exceptions.IllegalDataType.raise Else toDbValue = "NULL"
    ElseIf VBA.IsArray(value) Then
        If throwErrorIfIllegalDataType Then Call Exceptions.IllegalDataType.raise Else toDbValue = "NULL"
    ElseIf VBA.varType(value) = VBA.vbBoolean Then
        toDbValue = VBA.IIf(value, 1, 0)
    ElseIf VBA.varType(value) = VBA.vbString Then
        toDbValue = toDbString_typeAlreadyChecked(value)
    ElseIf f.Variables.isOfNumericType(value) Then
        toDbValue = toDbNumber_typeAlreadyChecked(value)
    ElseIf f.Dates.isTimeString(value) Then
        toDbValue = "'" & VBA.format(value, "HH:MM:SS") & "'"
    ElseIf VBA.isDate(value) Then
        toDbValue = "'" & VBA.format(value, "YYYY-MM-DD HH:MM:SS") & "'"
    End If

End Function

Public Function toDbString(value As Variant) As String
    Const METHOD_NAME As String = "toDbString"
    '------------------------------------------------------------------------------------------------------

    If Not f.Variables.isPrimitive(value) Then Call Exceptions.IllegalDataType.raise
    If f.Variables.isUndefined(value) Then
        toDbString = "NULL"
    Else
        toDbString = toDbString_typeAlreadyChecked(value)
    End If

End Function

Private Function toDbString_typeAlreadyChecked(ByVal value As Variant) As String
    Const METHOD_NAME As String = "toDbString_typeAlreadyChecked"
    '------------------------------------------------------------------------------------------------------
    Const REGEX_NON_PRINTABLE_CHARACTERS As String = "[\x00-\x1F]"
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------
    Dim withoutNonPrintableCharacters As String
    '------------------------------------------------------------------------------------------------------
    
    If objRegex Is Nothing Then Set objRegex = f.regex.Create(REGEX_NON_PRINTABLE_CHARACTERS)
    
    withoutNonPrintableCharacters = objRegex.replace(VBA.CStr(value), vbNullString)
    toDbString_typeAlreadyChecked = "'" & VBA.replace(withoutNonPrintableCharacters, "'", "''") & "'"

End Function

'Private Function hasAnyBlankCharacters(ByVal value As Variant) As Boolean
'    Const METHOD_NAME As String = "hasAnyBlankCharacters"
'    '------------------------------------------------------------------------------------------------------
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'    If VBA.InStr(1, value, VBA.Chr(10)) Then
'        hasAnyBlankCharacters = True
'    ElseIf VBA.InStr(1, value, VBA.Chr(13)) Then
'        hasAnyBlankCharacters = True
'    ElseIf VBA.InStr(1, value, VBA.Chr(9)) Then
'        hasAnyBlankCharacters = True
'    End If
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
''----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("value", value))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function

'Private Function toDbString_replaceBlankCharacters(ByVal value As Variant) As String
'    Const METHOD_NAME As String = "toDbString_replaceBlankCharacters"
'    '------------------------------------------------------------------------------------------------------
'    Dim i As Long
'    Dim asc As Integer
'    Dim char As String * 1
'    Dim substr As String
'    Dim arr() As String
'    '------------------------------------------------------------------------------------------------------
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'    For i = 1 To VBA.Len(value)
'        char = VBA.mid$(value, i, 1)
'        asc = VBA.asc(char)
'        If isAsciiForBlank(asc) Then
'            If VBA.Len(substr) Then Call F.arrays.addEntry(arr, "'" & substr & "'")
'            Call F.arrays.addEntry(arr, "CHAR(" & asc & ")")
'            substr = vbNullString
'        Else
'            substr = substr & char
'        End If
'    Next i
'    If VBA.Len(substr) Then Call F.arrays.addEntry(arr, "'" & substr & "'")
'
'    toDbString_replaceBlankCharacters = VBA.join(arr, " + ")
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
''----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("value", value))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function

'Private Function isAsciiForBlank(asc As Integer) As Boolean
'    If asc = 13 Then
'        isAsciiForBlank = True
'    ElseIf asc = 10 Then
'        isAsciiForBlank = True
'    ElseIf asc = 9 Then
'        isAsciiForBlank = True
'    End If
'End Function

Public Function toDbNumber(value As Variant) As String
    Const METHOD_NAME As String = "toDbNumber"
    '------------------------------------------------------------------------------------------------------
    Dim stringified As String
    '------------------------------------------------------------------------------------------------------

    If Not f.Variables.isPrimitive(value) Then Call Exceptions.IllegalDataType.raise
    
    If f.Variables.isUndefined(value) Then
        toDbNumber = "NULL"
    ElseIf f.Variables.isOfNumericType(value) Then
        toDbNumber = toDbNumber_typeAlreadyChecked(value)
    Else
        stringified = VBA.CStr(value)
        If isNumeric(stringified) Then
            toDbNumber = toDbNumber_typeAlreadyChecked(value)
        Else
            Call Exceptions.IllegalDataType.raise
        End If
    End If

End Function

Private Function toDbNumber_typeAlreadyChecked(value As Variant) As String
    Const METHOD_NAME As String = "toDbNumber_typeAlreadyChecked"
    '------------------------------------------------------------------------------------------------------
    toDbNumber_typeAlreadyChecked = trimLeft(trimRight(VBA.replace(value, ",", ".")))
End Function



'----------------------------------------------------------------------------------------------------------
' Description:          Function returns string representation of the given value.
' Returned value:       Function returns the following values depending on the type of the input value:
'                           * String - exactly the same value,
'                           * Boolean - text 'False' or 'True', depending on the value,
'                           * numeric data types - value converted to string (dot as a decimal separator),
'                           * Date (with time only) - string representation of the given time in format
'                             HH:MM:SS,
'                           * Date (full) - string representation of the given date and time in format
'                             YYYY-MM-DD HH:MM:SS,
'                           * undefined value (Empty / Null / Missing / Nothing):
'                                   - empty string in normal mode,
'                                   - 'NULL' in db mode,
'                                   - text #Empty / #Null / #Missing / #Nothing in XML mode,
'                           * Array - string representation of array (including information about each
'                             value stored in this array,
'                           * Collection - string representation of Collection (including information
'                             about each value stored in this Collection),
'                           * Dictionary - string representation of Dictionary (including information
'                             about each key-value pair stored in this Dictionary),
'                           * non-container Object:
'                                   - if object has its own [toString] method, result of this method
'                                     is returned,
'                                   - description of this object including data type and pointer.
'
' Parameters:
'   * value             Value that string representation is to be returned.
'   * stringifyMode     (Optional) It determines what mode should be used when returning string
'                       representation of the given value. Default value is StringifyMode_Normal.
'----------------------------------------------------------------------------------------------------------
Public Function stringify(value As Variant, _
                              Optional stringifyMode As StringifyModeEnum = StringifyMode_Normal, _
                              Optional includeNonPrimitiveDetails As Boolean = False) As String
    Const METHOD_NAME As String = "stringify"
    '------------------------------------------------------------------------------------------------------

    If f.Variables.isUndefined(value, False, False, False) Then
        stringify = stringifyUndefined(value, stringifyMode)
    ElseIf VBA.IsObject(value) Then
        stringify = stringifyObject(f.cast.toObject(value), stringifyMode, includeNonPrimitiveDetails)
    ElseIf f.Variables.isOfNumericType(value) Then
        stringify = stringifyNumericValue(value, stringifyMode)
    ElseIf VBA.IsArray(value) Then
        If includeNonPrimitiveDetails Then
            stringify = f.arrays.toString(value, stringifyMode)
        Else
            stringify = f.arrays.getDescription(value)
        End If
    Else
        If stringifyMode = StringifyMode_Db Then
            stringify = toDbValue(value)
        ElseIf f.Dates.isTimeString(value) Then
            stringify = VBA.format(value, "HH:MM:SS")
        ElseIf VBA.isDate(value) And Not VBA.varType(value) = VBA.vbString Then
            stringify = VBA.format(value, "YYYY-MM-DD HH:MM:SS")
        Else
            stringify = VBA.CStr(value)
        End If
    End If

End Function

Private Function stringifyUndefined(value As Variant, stringifyMode As StringifyModeEnum) As String
    Const METHOD_NAME As String = "stringifyUndefined"
    '------------------------------------------------------------------------------------------------------

    If stringifyMode = StringifyMode_Normal Then
        stringifyUndefined = vbNullString
    ElseIf stringifyMode = StringifyMode_Db Then
        stringifyUndefined = DB_NULL
    ElseIf stringifyMode = StringifyMode_Xml Then
        If VBA.IsObject(value) Then
            stringifyUndefined = XML_NOTHING
        ElseIf VBA.IsArray(value) Then
            stringifyUndefined = XML_ARRAY
        ElseIf VBA.IsEmpty(value) Then
            stringifyUndefined = XML_EMPTY
        ElseIf VBA.IsMissing(value) Then
            stringifyUndefined = XML_MISSING
        ElseIf VBA.IsNull(value) Then
            stringifyUndefined = XML_NULL
        End If
    End If

End Function

Private Function stringifyObject(value As Object, stringifyMode As StringifyModeEnum, _
                              Optional includeNonPrimitiveDetails As Boolean = False) As String
    Const METHOD_NAME As String = "stringifyObject"
    '------------------------------------------------------------------------------------------------------
    Const OBJECT_TAG = "[Data type: {0}; Pointer: {1}]"
    '------------------------------------------------------------------------------------------------------

    If value Is Nothing Then
        stringifyObject = VBA.IIf(stringifyMode = StringifyMode_Xml, XML_NOTHING, vbNullString)
    ElseIf Not includeNonPrimitiveDetails Then
        stringifyObject = f.strings.format(OBJECT_TAG, VBA.TypeName(value), VBA.ObjPtr(value))
    ElseIf TypeOf value Is VBA.Collection Then
        stringifyObject = f.Collections.toString(value, stringifyMode:=stringifyMode)
    ElseIf f.dictionaries.isDictionary(value) Then
        stringifyObject = f.dictionaries.toString(value, DictPart_KeyAndValue, stringifyMode:=stringifyMode)
    ElseIf TypeOf value Is Excel.Range Then
        stringifyObject = f.ranges.toString(value)
    Else
        stringifyObject = stringifyNotContainerObject(value, stringifyMode)
    End If

End Function

Private Function stringifyNotContainerObject(value As Object, stringifyMode As StringifyModeEnum) As String
    Const METHOD_NAME As String = "stringifyNotContainerObject"
    '------------------------------------------------------------------------------------------------------
    Const OBJECT_TAG = "[Object | DataType: {0} | Pointer: {1}]"
    Const XML_OBJECT_TAG = "<object><type>{0}</type><pointer>{1}</pointer></object>"
    '------------------------------------------------------------------------------------------------------
    Dim pattern As String
    Dim errNumber As Long, errDescription As String
    '------------------------------------------------------------------------------------------------------

    
    On Error Resume Next
    If stringifyMode = StringifyMode_Xml Then
        stringifyNotContainerObject = value.toXml
    Else
        stringifyNotContainerObject = value.toString
    End If
    errNumber = VBA.err.number
    errDescription = VBA.err.description


    If Not DEV_MODE Then On Error GoTo ErrHandler
    

    If errNumber = Exceptions.MethodNotSupported.getNumber Then
        pattern = OBJECT_TAG
        stringifyNotContainerObject = f.strings.format(pattern, VBA.TypeName(value), VBA.CLng(VBA.ObjPtr(value)))
        errNumber = 0
    ElseIf errNumber > 0 Then
        Call VBA.err.raise(errNumber, description:=errDescription)
    End If
    
    
    If stringifyMode = StringifyMode_Db Then
        stringifyNotContainerObject = toDbString_typeAlreadyChecked(stringifyNotContainerObject)
    End If

'==========================================================================================================
ExitPoint:
    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
                                "value; stringifyMode; pattern", value, stringifyMode, pattern)
    GoTo ExitPoint

End Function

Private Function stringifyNumericValue(value As Variant, stringifyMode As StringifyModeEnum) As String
    Const METHOD_NAME As String = "stringifyNumericValue"
    '------------------------------------------------------------------------------------------------------

    If stringifyMode = StringifyMode_Db Then
        stringifyNumericValue = toDbNumber(value)
    Else
        stringifyNumericValue = VBA.CStr(value)
    End If

End Function



'**********************************************************************************************************
' Name:                 substring
' Author:               mielk | 2011-07-20
'
' Comment:              Function to return the substring of a source string between two given
'                       substring delimiters.
'
'                       Technically, the function looks for the position of the first occurrence of the
'                       given left delimiter. Then, it looks for the first occurrence of the given right
'                       delimiter after that position. Finally it returns substring between those two
'                       positions.
'
' Parameters:
'   sourceText          Base string.
'
'   leftSubstr          Left delimiter. The result string will start after the first occurrence of this
'                       substring in the source text.
'                       If this substring is empty, the result string will start from the beginning of the
'                       source text.
'
'   rightSubstr         Right delimiter. The result string will end before first occurence of this string
'                       after the position of the leftSubstr.
'                       If this substring is empty, the result string will contain all the characters to
'                       the end of the source text.

'   isCaseSensitive     Optional parameter of Boolean type.
'                       It determines if text searching is case sensitive.
'
'                       If this value is set to True, searching is case sensitive - a letter in lowercase
'                       is treated as different than the same letter in uppercase (i.e. a ? A).
'
'                       If this value is set to False, it doesn't matter if a letter is in lowercase or
'                       in uppercase, since both of them are considered as the same character (i.e. a = A).
'
'                       Default value of this parameter is True.
'
'
' Returns:
'   String              Substring between first occurencies of [leftSubstr] and [rightSubstr].
'
'                       If there is either no [leftSubstr] in the source text or no [rightSubstr] after
'                       this [leftSubstr], then empty string is returned.
'
'                       If [leftSubstr] is empty string, the result substring starts from the beginning of
'                       the base string.
'
'                       If [rightSubstr] is empty string, the result substring will contain all the
'                       characters between the first occurrence of the [leftSubstr] and the end of
'                       the source text.
'
'                       Examples:
'                       ---------------------------------------------------------------------------------
'                                   Function with parameters                |     Result
'                               -------------------------------------------------------------
'                               substring("abc", "a", "c", True)            |       b
'                               substring("abC", "a", "c", True)            |
'                               substring("abC", "a", "c", False)           |       b
'                               substring("abc", "", "c", True)             |       ab
'                               substring("abc", "a", "", True)             |       bc
'                               substring("abC", "", "c", False)            |       ab
'                               substring("abC", "", "c", True)             |
'                               substring("[value]", "[", "]", True)        |       value
'                               substring("[value]", "[", "]", False)       |       value
'                               substring("aaabbb", "a", "b", True)         |       aa
'                               substring("aaabbb", "a", "a", True)         |
'                               substring("aaabbb", "", "ab", True)         |       aa
'                               substring("aaabbb", "", "a", True)          |
'                               substring("aaabbb", "b", "", True)          |       bb
'                               substring("one two three", " ", " ", True)  |       two
'                               substring("one two three", " ", "", True)   |       two three
'
'
' --- Changes log -----------------------------------------------------------------------------------------
' 2011-07-20    mielk       Function created.
'**********************************************************************************************************
Public Function substring(sourceText As String, leftSubstr As String, rightSubstr As String, _
                                                    Optional isCaseSensitive As Boolean = True) As String
    Const METHOD_NAME As String = "substring"
    '------------------------------------------------------------------------------------------------------
    Dim startPosition As Long
    Dim endPosition As Long
    Dim compareMethod As VBA.VbCompareMethod
    Dim resultLength As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    compareMethod = VBA.IIf(isCaseSensitive, VBA.vbBinaryCompare, VBA.vbTextCompare)

    If VBA.Len(leftSubstr) Then
        startPosition = VBA.InStr(1, sourceText, leftSubstr, compareMethod)
    Else
        startPosition = 1
    End If

    'No point to look for [endPosition] if start position is 0, result will be empty string anyway.
    If startPosition > 0 Then
        startPosition = startPosition + VBA.Len(leftSubstr)
        If VBA.Len(rightSubstr) Then
            endPosition = VBA.InStr(startPosition, sourceText, rightSubstr, compareMethod)
        Else
            endPosition = VBA.Len(sourceText) + 1
        End If
    End If

    If startPosition > 0 And endPosition > 0 Then
        resultLength = endPosition - startPosition
        substring = VBA.mid$(sourceText, startPosition, resultLength)
    End If

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then
        Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
                    "sourceText; leftSubstr; rightSubstr; isCaseSensitive; startPosition; endPosition; compareMethod; resultLength", _
                     sourceText, leftSubstr, rightSubstr, isCaseSensitive, startPosition, endPosition, compareMethod, resultLength)
    GoTo ExitPoint

End Function





'----------------------------------------------------------------------------------------------------------
' Description:          Function takes dynamic number of arguments and returns the first of them that is
'                       a primitive value other than empty string.
'
' Returned value:       Function checks all the given arguments and returns the first one that has its
'                       string representation containing at least one character. Arguments are being
'                       checked in the same order they were passed when invoking the function.

'                       If none of arguments passed to the function meets the requirements described above,
'                       function returns vbNullString.
'
' Parameters:
'   * strings           strings to be checked. They are checked in the same order they were passed when
'                       invoking the function.
'----------------------------------------------------------------------------------------------------------
Public Function coalesce(ParamArray strings() As Variant) As String
    Const METHOD_NAME As String = "coalesce"
    '------------------------------------------------------------------------------------------------------
    Dim varString As Variant
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For Each varString In strings
        If isDefinedPrimitiveVariable(varString) Then
            If VBA.Len(varString) Then
                coalesce = VBA.CStr(varString)
                Exit For
            End If
        End If
    Next varString

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
                    "strings; varString", f.arrays.fromParamArray(strings), varString)
    GoTo ExitPoint

End Function


Private Function isDefinedPrimitiveVariable(value As Variant) As Boolean
    Const METHOD_NAME As String = "isDefinedPrimitiveVariable"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If f.Variables.isUndefined(value) Then
        isDefinedPrimitiveVariable = False
    ElseIf VBA.IsArray(value) Then
        isDefinedPrimitiveVariable = False
    ElseIf VBA.IsObject(value) Then
        isDefinedPrimitiveVariable = False
    Else
        isDefinedPrimitiveVariable = True
    End If

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "value", value)
    GoTo ExitPoint

End Function



'#todo - zamieni� na regex
Public Function onlyDigits(ByVal text As Variant, _
              Optional ByVal leaveSpecialChars As Boolean = True) As String
    Const METHOD_NAME As String = "onlyDigits"
    '------------------------------------------------------------------------------------------------------
    Const MINUS As String = "-"
    Const COMA As String = ","
    Const DOT As String = "."
    '------------------------------------------------------------------------------------------------------
    Dim sText As String
    Dim iChar As Integer
    Dim sChar As String
    Dim iAsciiCode As Integer
    Dim bHasComa As Boolean
    '------------------------------------------------------------------------------------------------------


    On Error GoTo IllegalTypeException
    sText = VBA.CStr(text)

    For iChar = 1 To VBA.Len(sText)
        sChar = VBA.mid$(sText, iChar, 1)
        If isDigit(sChar) Then
            onlyDigits = onlyDigits & sChar
        Else
            If leaveSpecialChars Then
                If sChar = MINUS Then
                    If VBA.Len(onlyDigits) = 0 Then
                        If iChar < VBA.Len(sText) Then
                            If isDigit(VBA.mid$(sText, iChar + 1, 1)) Then
                                onlyDigits = MINUS
                            End If
                        End If
                    End If
                ElseIf Not bHasComa And (sChar = COMA Or sChar = DOT) Then
                    If VBA.Len(onlyDigits) Then
                        If VBA.right$(onlyDigits, 1) <> MINUS Then
                            onlyDigits = onlyDigits & Application.DecimalSeparator
                            bHasComa = True
                        End If
                    End If
                End If
            End If
        End If
    Next iChar


'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
IllegalTypeException:
    GoTo ExitPoint

End Function

Public Function isDigit(ByVal char As String) As Boolean
    Dim iAsciiCode As Integer
    '------------------------------------------------------------------------------------------------------
    iAsciiCode = VBA.asc(char)
    isDigit = (iAsciiCode >= 48 And iAsciiCode <= 57)
End Function









'----------------------------------------------------------------------------------------------------------
' Description:          Function removes all blank characters at the beginning of the given string.
'
' Returned value:       Copy of the given string without blank characters at the beginning.
'
' Parameters:
'   * text              Text to be trimmed.
'----------------------------------------------------------------------------------------------------------
Public Function trimLeft(text As String) As String
    Const METHOD_NAME As String = "trimLeft"
    '------------------------------------------------------------------------------------------------------
    Const REGEX_PATTERN As String = "^\s+"
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If objRegex Is Nothing Then
        Set objRegex = f.regex.Create(REGEX_PATTERN)
    End If
    trimLeft = objRegex.replace(text, vbNullString)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "text", text)
    GoTo ExitPoint

End Function





'----------------------------------------------------------------------------------------------------------
' Description:          Function removes all blank characters at the end of the given string.
'
' Returned value:       Copy of the given string without blank characters at the end.
'
' Parameters:
'   * text              Text to be trimmed.
'----------------------------------------------------------------------------------------------------------
Public Function trimRight(text As String) As String
    Const METHOD_NAME As String = "trimRight"
    '------------------------------------------------------------------------------------------------------
    Const REGEX_PATTERN As String = "\s+$"
    Static objRegex As Object
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If objRegex Is Nothing Then
        Set objRegex = f.regex.Create(REGEX_PATTERN)
    End If
    trimRight = objRegex.replace(text, vbNullString)

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "text", text)
    GoTo ExitPoint

End Function






'----------------------------------------------------------------------------------------------------------
' Description:          Function removes all redundant blank characters in the middle of the given string.
'
' Returned value:       Copy of the given string without duplicated blank characters in the middle.
'                       If there are groups of more than one blank characters in the original string,
'                       the result string will contain only first of those blank characters.
'
' Parameters:
'   * text              Text to be trimmed.
'----------------------------------------------------------------------------------------------------------
Public Function trimInside(text As String) As String
    Const METHOD_NAME As String = "trimInside"
    '------------------------------------------------------------------------------------------------------
    Const REGEX_PATTERN_NEW_LINE As String = "\b(\s+(\r|\f|\n)|(\r|\f|\n)\s+)\b"
    Const REGEX_PATTERN_TAB As String = "\b\s*(\t)\s*\b"
    Const REGEX_PATTERN_SPACES As String = "\b  *\b"
    '------------------------------------------------------------------------------------------------------
    Static objRegexNewLine As Object
    Static objRegexTab As Object
    Static objRegexSpace As Object
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If objRegexNewLine Is Nothing Then Set objRegexNewLine = f.regex.Create(REGEX_PATTERN_NEW_LINE)
    If objRegexTab Is Nothing Then Set objRegexTab = f.regex.Create(REGEX_PATTERN_TAB)
    If objRegexSpace Is Nothing Then Set objRegexSpace = f.regex.Create(REGEX_PATTERN_SPACES)

    trimInside = objRegexNewLine.replace(text, vbCrLf)
    trimInside = objRegexTab.replace(trimInside, vbTab)
    trimInside = objRegexSpace.replace(trimInside, " ")

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "text", text)
    GoTo ExitPoint

End Function





'----------------------------------------------------------------------------------------------------------
' Description:          Function removes all redundant blank characters at the beginning, in the middle and
'                       at the end of the given string.
'
' Returned value:       Copy of the given string without redundant blank characters.
'                       The following blank characters are considered to be redundant:
'                        * all blank characters at the beginning,
'                        * all blank characters at the end,
'                        * all blank characters in the middle that are preceded by other blank character.
'
' Parameters:
'   * text              Text to be trimmed.
'----------------------------------------------------------------------------------------------------------
Public Function trimFull(text As String) As String
    Const METHOD_NAME As String = "trimFull"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    trimFull = trimLeft(trimInside(trimRight(text)))

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, "text", text)
    GoTo ExitPoint

End Function




Public Function addLeadingZeros(ByVal value As Variant, Optional ByVal strLength As Byte = 2) As String
    Dim iMissingChars As Integer
    Dim lNumber As Long
    '------------------------------------------------------------------------------------------------------

    If VBA.isNumeric(value) Then
        lNumber = VBA.Int(VBA.Abs(value))
        iMissingChars = strLength - VBA.Len(VBA.CStr(lNumber))
        If iMissingChars > 0 Then
            addLeadingZeros = VBA.String(iMissingChars, "0") & lNumber
        Else
            addLeadingZeros = lNumber
        End If
    Else
        On Error GoTo IllegalTypeException
        addLeadingZeros = VBA.CStr(value)
    End If

'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
IllegalTypeException:
    GoTo ExitPoint

End Function


Public Function cutFromEnd(baseText As String, textToCut As String, Optional caseSensitive As Boolean = False) As String
    If VBA.Len(textToCut) = 0 Or VBA.Len(baseText) = 0 Then
        cutFromEnd = baseText
    ElseIf f.strings.compareStrings(VBA.right$(baseText, VBA.Len(textToCut)), textToCut, caseSensitive) Then
        cutFromEnd = VBA.left$(baseText, VBA.Len(baseText) - VBA.Len(textToCut))
    Else
        cutFromEnd = baseText
    End If
End Function




'----------------------------------------------------------------------------------------------------------
' Description:          Function counts and returns how many times the given substring [substr] appears in
'                       the given [sourceText].
'
' Returned value:       The number of occurrences of the given substring [substr] in the given source text.
'
' Parameters:
'   * sourceText        Original string.
'   * substr            Substring that occurrences are counted.
'   * isCaseSensitive   (Optional) It determines if searching for substring is case-sensitive.
'----------------------------------------------------------------------------------------------------------
Public Function countOccurrences(sourceText As String, substr As String, _
                                                     Optional isCaseSensitive As Boolean = True) As Long
    Const METHOD_NAME As String = "countOccurrences"
    '------------------------------------------------------------------------------------------------------
    Dim position As Integer
    Dim uCompareMethod As VBA.VbCompareMethod
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    uCompareMethod = VBA.IIf(isCaseSensitive, VBA.vbBinaryCompare, VBA.vbTextCompare)
    If VBA.Len(sourceText) > 0 And VBA.Len(substr) > 0 Then
        Do
            position = VBA.InStr(position + 1, sourceText, substr, uCompareMethod)
            If position Then countOccurrences = countOccurrences + 1 Else Exit Do
        Loop
    End If

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Function

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = err.number: errDescription = err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
                                "sourceText; substr; isCaseSensitive", _
                                 sourceText, substr, isCaseSensitive)
    GoTo ExitPoint

End Function


Public Function latinize(ByVal text As String) As String
    Dim str As String:              str = text
    '------------------------------------------------------------------------------------------------------
    str = VBA.replace(str, "�", "a")
    str = VBA.replace(str, "�", "A")
    str = VBA.replace(str, "�", "c")
    str = VBA.replace(str, "�", "C")
    str = VBA.replace(str, "�", "e")
    str = VBA.replace(str, "�", "E")
    str = VBA.replace(str, "�", "l")
    str = VBA.replace(str, "�", "L")
    str = VBA.replace(str, "�", "n")
    str = VBA.replace(str, "�", "N")
    str = VBA.replace(str, "�", "o")
    str = VBA.replace(str, "�", "O")
    str = VBA.replace(str, "�", "s")
    str = VBA.replace(str, "�", "S")
    str = VBA.replace(str, "�", "z")
    str = VBA.replace(str, "�", "Z")
    str = VBA.replace(str, "�", "z")
    str = VBA.replace(str, "�", "Z")
    latinize = str
End Function




Public Function removeNonWordCharacters(text As String) As String
    Const REGEX_PATTERN As String = "[^\w]"
    '------------------------------------------------------------------------------------------------------
    removeNonWordCharacters = f.regex.replace(text, REGEX_PATTERN, vbNullString)
End Function

Public Function toCharactersArray(text As String) As Variant
    Dim arr As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    If VBA.Len(text) Then
        ReDim arr(1 To VBA.Len(text))
        For i = 1 To VBA.Len(text)
            arr(i) = VBA.mid$(text, i, 1)
        Next i
        toCharactersArray = arr
    Else
        toCharactersArray = f.arrays.createEmpty1DArray
    End If
End Function





Public Function replaceByDictionary(baseString As String, replacements As Scripting.Dictionary) As String
    Dim key As Variant
    Dim value As String
    '------------------------------------------------------------------------------------------------------
    replaceByDictionary = baseString
    If f.dictionaries.isNonEmptyDictionary(replacements) Then
        For Each key In replacements.Keys
            value = replacements.Item(key)
            replaceByDictionary = VBA.replace(replaceByDictionary, VBA.CStr(key), value)
        Next key
    End If
End Function






'
'
'
''----------------------------------------------------------------------------------------------------------
'' Description:          Function counts and returns how many lines does the given text have.
''
'' Returned value:       The number of lines in the given text.
''
'' Parameters:
''   * text              Text which lines are counted.
''----------------------------------------------------------------------------------------------------------
'Public Function countTextRows(ByVal text As String) As Long
'    Const METHOD_NAME As String = "countTextRows"
'    '------------------------------------------------------------------------------------------------------
'    Dim carriageReturns As Integer
'    Dim lineFeeds As Integer
'    '------------------------------------------------------------------------------------------------------
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'    text = VBA.Replace(trimRight(text), vbCrLf, vbCr)
'    If VBA.Len(text) Then
'        carriageReturns = countOccurrences(text, VBA.Chr$(13), True)
'        lineFeeds = countOccurrences(text, VBA.Chr$(10), True)
'        countTextRows = carriageReturns + lineFeeds + 1
'    End If
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
''----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("text", text))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function
'
'
'
'
''----------------------------------------------------------------------------------------------------------
'' Description:          Wrapping function to put the given string in quotes.
'' Returned value:       The original text inserted into quotes.
''
'' Parameters:
''   * str               Text to be inserted into quotes.
''----------------------------------------------------------------------------------------------------------
'Public Function putInQuotes(str As String) As String
'    putInQuotes = VBA.Chr(34) & str & VBA.Chr(34)
'End Function
'
'
'
'
'
''----------------------------------------------------------------------------------------------------------
'' Description:          Wrapping function to check if the given string [baseString] contains another string
''                       [lookFor].
'' Returned value:       True - if the given string [baseString] contains another [lookFor].
''
'' Parameters:
''   * baseString        String that is searched for substring [lookFor].
''   * lookFor           Substring to be looked for.
''   * isCaseSensitive   (Optional) It defines if searching is case-sensitive.
''----------------------------------------------------------------------------------------------------------
'Public Function contains(baseString As String, lookFor As String, _
'                                                    Optional isCaseSensitive As Boolean = False) As Boolean
'    Const METHOD_NAME As String = "contains"
'    '------------------------------------------------------------------------------------------------------
'    Dim uCompareMethod As VBA.VbCompareMethod
'    '------------------------------------------------------------------------------------------------------
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'    uCompareMethod = VBA.IIf(isCaseSensitive, VBA.vbBinaryCompare, VBA.vbTextCompare)
'    contains = (VBA.InStr(1, baseString, lookFor, uCompareMethod) > 0)
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
''----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("baseString", baseString), _
'                    KeyValue("lookFor", lookFor), KeyValue("isCaseSensitive", isCaseSensitive))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function
'
'
'
'Public Function isCorrectUuid(value As String) As Boolean
'    Const METHOD_NAME As String = "isCorrectUuid"
'    '------------------------------------------------------------------------------------------------------
'    Const UUID_REGEX_PATTERN As String = _
'               "^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[1-5][0-9A-Fa-f]{3}-[89aAbB][0-9A-Fa-f]{3}-[0-9A-Fa-f]{12}$"
'    Static objRegex As Object
'    '------------------------------------------------------------------------------------------------------
'
'    If Not DEV_MODE Then On Error GoTo ErrHandler
'
'    If objRegex Is Nothing Then
'        Set objRegex = F.regex.Create(UUID_REGEX_PATTERN)
'    End If
'    isCorrectUuid = objRegex.test(value)
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
''----------------------------------------------------------------------------------------------------------
'ErrHandler:
'    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
'    Set err = VBA.err: errNumber = err.number: errDescription = err.description
'    Set params = F.Dictionaries.createWithItems(False, KeyValue("value", value))
'    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
'    Call VBA.err.raise(errNumber, source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
'
'End Function
'
'
'
'
'
'
'
'
'
''***** WITHOUT UNIT TESTS *****

''
''
'''*** TODO: unit testy ***
''Public Function getSlashSeparatedString(ParamArray substrings() As Variant) As String
''    getSlashSeparatedString = VBA.Join(substrings, "\")
''End Function
''
''Public Function separatorLine(Optional includeNewLine As Boolean = True, Optional character As String = "-") As String
''    separatorLine = VBA.String$(50, character) & VBA.IIf(includeNewLine, VBA.Chr$(10), vbNullString)
''End Function
''
''Public Function addLeadingZeros(ByVal value As Variant, Optional ByVal desiredLength As Byte = 2) As String
''    Dim missingCharsCounter As Integer
''    Dim number As Long
''    '------------------------------------------------------------------------------------------------------
''
''    Call throwExceptionIfNotIsNumeric(value)
''
''    number = VBA.Int(VBA.Abs(value))
''    missingCharsCounter = desiredLength - VBA.Len(VBA.CStr(number))
''    If missingCharsCounter > 0 Then
''        addLeadingZeros = VBA.String(missingCharsCounter, "0") & number
''    Else
''        addLeadingZeros = number
''    End If
''
''End Function
'
'
'Public Function removeSubstrings(baseString As String, substrings As Variant) As String
'    Dim substring As Variant
'    '------------------------------------------------------------------------------------------------------
'    removeSubstrings = baseString
'    If F.arrays.isDefined(substrings) Then
'        For Each substring In substrings
'            removeSubstrings = VBA.Replace(removeSubstrings, VBA.CStr(substring), vbNullString)
'        Next substring
'    End If
'End Function
'
''Public Function getStringsSimilarity(str1 As String, str2 As String) As Double
''    'VBA.InStr(1, varName, searchName)
''End Function
'
'Public Function getRegexLastMatch(matches As Object) As Object
'    Dim i As Long
'    Dim item As Variant
'    Dim count As Long
'    '------------------------------------------------------------------------------------------------------
'    For Each item In matches
'        count = count + 1
'        If count = matches.count Then
'            Set getRegexLastMatch = item
'            Exit For
'        End If
'    Next item
'End Function
'
'Public Function onlyLetters(text As String) As String
'    Const REGEX_PATTERN As String = "[^a-zA-Z����󜟿��ʣ�ӌ��]"
'    '------------------------------------------------------------------------------------------------------
'    Static objRegex As Object
'    '------------------------------------------------------------------------------------------------------
'
'    If objRegex Is Nothing Then Set objRegex = F.regex.Create(REGEX_PATTERN)
'    onlyLetters = objRegex.Replace(text, vbNullString)
'
'End Function
'
'
'Public Function onlyDigits(ByVal text As Variant, _
'              Optional ByVal leaveSpecialChars As Boolean = True) As String
'    Const METHOD_NAME As String = "onlyDigits"
'    '------------------------------------------------------------------------------------------------------
'    Const MINUS As String = "-"
'    Const COMA As String = ","
'    Const DOT As String = "."
'    '------------------------------------------------------------------------------------------------------
'    Dim sText As String
'    Dim iChar As Integer
'    Dim sChar As String
'    Dim iAsciiCode As Integer
'    Dim bHasComa As Boolean
'    '------------------------------------------------------------------------------------------------------
'
'
'    On Error GoTo IllegalTypeException
'    sText = VBA.CStr(text)
'
'    For iChar = 1 To VBA.Len(sText)
'        sChar = VBA.mid$(sText, iChar, 1)
'        If isDigit(sChar) Then
'            onlyDigits = onlyDigits & sChar
'        Else
'            If leaveSpecialChars Then
'                If sChar = MINUS Then
'                    If VBA.Len(onlyDigits) = 0 Then
'                        If iChar < VBA.Len(sText) Then
'                            If isDigit(VBA.mid$(sText, iChar + 1, 1)) Then
'                                onlyDigits = MINUS
'                            End If
'                        End If
'                    End If
'                ElseIf Not bHasComa And (sChar = COMA Or sChar = DOT) Then
'                    If VBA.Len(onlyDigits) Then
'                        If VBA.right$(onlyDigits, 1) <> MINUS Then
'                            onlyDigits = onlyDigits & Application.DecimalSeparator
'                            bHasComa = True
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    Next iChar
'
'
''==========================================================================================================
'ExitPoint:
'    Exit Function
'
''----------------------------------------------------------------------------------------------------------
'IllegalTypeException:
'    GoTo ExitPoint
'
'End Function
'

'Public Function toSentenceCase(text As String) As String
'    toSentenceCase = VBA.UCase$(VBA.mid$(text, 1, 1)) & VBA.mid$(text, 2)
'End Function

'
'Public Function clearJsonString(value As String) As String
'    Dim pos As Long
'    '------------------------------------------------------------------------------------------------------
'    pos = VBA.InStr(1, value, "{")
'    If pos > 0 Then
'        clearJsonString = VBA.mid(value, pos)
'    End If
'End Function



