VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'sprawdziæ czym siê ró¿ni getColumnIndexByHeader od getColumnIndex (to samo dla Row)
'ogarn¹æ funkcje slice (wystêpuj¹ dwa razy)

Option Explicit

Private Const CLASS_NAME As String = "FArrays"
'----------------------------------------------------------------------------------------------------------
Private Const MULTIDIMENSIONAL_ARRAY_TAG As String = "[{0}-dimensional array]"
'----------------------------------------------------------------------------------------------------------



Public Function createEmpty1DArray() As Variant
    Dim arr() As Variant
    '----------------------------------------------------------------------------------------------------------
    createEmpty1DArray = arr
End Function


'[DIMENSIONS]
'----------------------------------------------------------------------------------------------------------
' Description:          Function checks if the given parameter [arr] is an array and has at least as many
'                       dimensions as the given parameter [dimensions].
' Returned value:       True - if the given parameter [arr] is a declared array and it has at least as many
'                       dimensions as the value of input parameter [dimension].
'                       False - in the following cases:
'                           1) input parameter [arr] is not an array,
'                           2) input parameter [arr] is an array but its size and dimensions has not been
'                               declared yet,
'                           3) input parameter [arr] is an array with already declared size and dimensions,
'                               but its number of dimensions is lower than the value of input parameter
'                               [dimension],
'                           4) value of input parameter [dimension] is equal or less than zero.
' Parameters:
'   * arr               Array to be checked.
'   * dimension         Index of dimension to be checked.
'   * dimension
'                       Index of dimension to be checked.
'----------------------------------------------------------------------------------------------------------
Public Function hasDimension(arr As Variant, dimension As Integer) As Boolean
    Const METHOD_NAME As String = "hasDimension"
    Dim errLocal As Object, errNumber As Long, errDescription As String
    '------------------------------------------------------------------------------------------------------
    Dim lowBound As Long
    Dim upBound As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If dimension > 0 Then
        On Error GoTo MissingDimensionException
        lowBound = LBound(arr, dimension)
        upBound = UBound(arr, dimension)
        
        If Not DEV_MODE Then On Error GoTo ErrHandler
        hasDimension = (upBound >= lowBound)
    End If

'==========================================================================================================
ExitPoint:

    '[Clean up]

    '[If any error was raised, re-throw it]
    If errNumber > 0 Then
        Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    
    Exit Function

    '----------------------------------------------------------------------------------------------------------
MissingDimensionException:
    GoTo ExitPoint
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim errParams As Scripting.Dictionary
    '- - - - - - - - - - - - - - - - - - -
    Set errLocal = VBA.err: errNumber = errLocal.number: errDescription = errLocal.description
    Set errParams = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
            KeyValue("dimension", dimension), KeyValue("lowBound", lowBound), KeyValue("upBound", upBound))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, errParams)
    GoTo ExitPoint

End Function








'----------------------------------------------------------------------------------------------------------
' Description:          Function checks if the given input parameter [arr] is an array and has its sizes
'                       and dimensions already declared.
' Returned value:       True - if the given parameter [arr] is an array and has its sizes and dimensions
'                       already declared.
'                       False - in the following cases:
'                           1) input parameter [arr] is not an array,
'                           2) input parameter [arr] is an array but its size and dimensions has not been
'                               declared yet.
' Parameters:
'   * arr               Array to be checked.
'----------------------------------------------------------------------------------------------------------
Public Function isDefined(arr As Variant) As Boolean
    Const METHOD_NAME As String = "isDefined"
    '------------------------------------------------------------------------------------------------------


    If Not DEV_MODE Then On Error GoTo ErrHandler
    If Not VBA.IsArray(arr) Then
        isDefined = False
    Else
        isDefined = hasDimension(arr, 1)
    End If
    

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function





'----------------------------------------------------------------------------------------------------------
' Description:          Function returns the number of dimensions for the given array [arr].
' Returned value:       * If the given parameter [arr] is not an array function returns -1.
'                       * If the given parameter [arr] is an array but it has not been defined yet (it has
'                         not its size and dimensions declared) function returns 0.
'                       * If the given parameter [arr] is an array and its size is already defined, function
'                         returns then number of dimensions of this array.
' Parameters:
'   * arr               Array for which the number of dimensions is returned.
'----------------------------------------------------------------------------------------------------------
Public Function countDimensions(arr As Variant) As Long
    Const METHOD_NAME As String = "countDimensions"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If VBA.IsArray(arr) Then
        Do
            If hasDimension(arr, countDimensions + 1) Then
                countDimensions = countDimensions + 1
            Else
                Exit Do
            End If
        Loop
    Else
        countDimensions = -1
    End If


'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function





'----------------------------------------------------------------------------------------------------------
' Description:          Wrapper function to check if the given parameter [value] is 1D array.
' Returned value:       True - if the given parameter [value] is a one-dimensional array.
'                       False - in the following cases:
'                           1) input parameter [value] is not an array,
'                           2) input parameter [value] is an array but its size and dimensions has not been
'                               declared yet,
'                           3) input parameter [value] is an array with more than one dimension.
' Parameters:
'   * value             Array to be checked.
'----------------------------------------------------------------------------------------------------------
Public Function is1D(value As Variant) As Boolean
    Const METHOD_NAME As String = "is1D"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If Not VBA.IsArray(value) Then
        is1D = False
    ElseIf countDimensions(value) = 1 Then
        is1D = True
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("value", value))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function checks if the given parameter [arr] is a dynamic array.
' Returned value:       True - if the given parameter [arr] is an array and it has been declared as dynamic
'                           (so its size can be changed during run-time).
'                       False - in the following cases:
'                           1) input parameter [arr] is not an array,
'                           2) input parameter [arr] is an array with fixed size.
' Parameters:
'   * arr               Array to be checked.
' Exceptions:
'   * WrongDimensionsNumberException        Thrown if the given parameter [arr] is an array with more than
'                                           3 dimensions.
    '----------------------------------------------------------------------------------------------------------
Public Function isDynamic(arr As Variant) As Boolean
    Const METHOD_NAME As String = "isDynamic"
    '------------------------------------------------------------------------------------------------------
    Dim dimensions As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If VBA.IsArray(arr) Then
        dimensions = countDimensions(arr)
        Select Case dimensions
            Case 0
                isDynamic = True
            Case 1
                On Error GoTo NotDynamicArray
                ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1))
                isDynamic = True
            Case 2
                On Error GoTo NotDynamicArray
                ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1), _
                                   LBound(arr, 2) To UBound(arr, 2))
                isDynamic = True
            Case 3
                On Error GoTo NotDynamicArray
                ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1), _
                                   LBound(arr, 2) To UBound(arr, 2), _
                                   LBound(arr, 3) To UBound(arr, 3))
                isDynamic = True
            Case Else
                Call Exceptions.InvalidDimension.raise
                
        End Select
    End If


'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
NotDynamicArray:
    GoTo ExitPoint

    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function





'[SIZE]
    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns the size of the given array for the given dimension.
' Returned value:       If the given parameter [arr] is an array and has at least as many dimensions as
'                           parameter [dimension], function returns the size of this array in this
'                           dimension.
'                       For dynamic arrays not defined yet, function returns 0.
'                       Function returns -1 in the following cases:
'                           1) input parameter [arr] is not an array,
'                           2) input parameter [dimensions] is less or equal to zero,
'                           3) input parameter [dimensions] is greater than the number of dimensions of
'                               the given array.
' Parameters:
'   * arr               Array for which size is to be returned.
'   * dimension         Index of dimension for which size is to be returned.
    '----------------------------------------------------------------------------------------------------------
Public Function getSize(arr As Variant, Optional dimension As Integer = 1) As Long
    Const METHOD_NAME As String = "getSize"
    '------------------------------------------------------------------------------------------------------
    Dim dimensions As Integer
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler


    If dimension <= 0 Then
        getSize = -1
    ElseIf VBA.IsArray(arr) Then
        dimensions = countDimensions(arr)
        If dimensions = 0 Then
            getSize = 0
        ElseIf dimension > dimensions Then
            getSize = -1
        Else
            getSize = UBound(arr, dimension) - LBound(arr, dimension) + 1
        End If
    Else
        getSize = -1
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
                                                                        KeyValue("dimension", dimension))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function

'For backward compatibility.
Public Function arraySize(arr As Variant, Optional dimension As Integer = 1) As Long
    arraySize = getSize(arr, dimension)
End Function




    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns the number of items in the given array.
' Returned value:       If the given parameter [arr] is an array with its size defined, function returns
'                           the number of items in this array (size of reach dimension multiplied by each
'                           other).
'                       If the given parameter [arr] is a dynamic array and its size has not been defined
'                           yet, function returns 0.
'                       If the given parameter [arr] is not an array, function returns -1.
' Parameters:
'   * arr               Array for which the number of items is returned.
    '----------------------------------------------------------------------------------------------------------
Public Function countItems(arr As Variant) As Long
    Const METHOD_NAME As String = "countItems"
    '------------------------------------------------------------------------------------------------------
    Dim dimensions As Integer
    Dim i As Integer
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If VBA.IsArray(arr) Then
        dimensions = countDimensions(arr)
        '[Must be initialized with 1 before multiplying]
        countItems = VBA.IIf(dimensions > 0, 1, 0)
        For i = 1 To dimensions
            countItems = countItems * getSize(arr, i)
        Next i
    Else
        countItems = -1
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function compares the size of two arrays and returns True if they are equal.
' Returned value:       Function returns True if:
'                           * both parameters [arr1] and [arr2] are arrays and they have the same number of
'                             dimensions and the same size in each dimension,
'                           * both parameters [arr1] and [arr2] are dynamic arrays without their sizes
'                             defined yet.
'                       Function returns False if:
'                           * at least one of the given parameters is not an array,
'                           * given parameters are arrays but they have different number of dimensions,
'                           * given parameters are arrays with the same number of dimensions, but they have
'                             different size in any of dimension,
'                           * given parameters are arrays with the same number of dimensions and the equal
'                             size in each dimension, but they have different indexing and optional
'                             parameter [allowDifferentIndexing] is set to False.
' Parameters:
'   * arr1, arr2        Arrays to be compared.
'   * allowDifferentIndexing
'                       (Optional) If set to True, compared arrays can have different indexing, dimension.e. for
'                           arrays (1 To 2, 1 To 3) and (0 To 1, 0 To 2) function would return True,
'                           because both are 2x3 arrays.
'                           If set to False, indexing for each dimension in each array must start from the
'                           same value for the function to return True.
    '----------------------------------------------------------------------------------------------------------
Public Function compareSizes(arr1 As Variant, arr2 As Variant, _
                                            Optional allowDifferentIndexing As Boolean = True) As Boolean
    Const METHOD_NAME As String = "compareSizes"
    '------------------------------------------------------------------------------------------------------
    Dim dimensions As Integer
    Dim dimension As Integer
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(arr1) Then
        compareSizes = False
    ElseIf Not VBA.IsArray(arr2) Then
        compareSizes = False
    ElseIf compareDimensions(arr1, arr2) Then
        compareSizes = True
        dimensions = countDimensions(arr1)
        For dimension = 1 To dimensions
            If Not compareDimensionsSize(arr1, arr2, dimension, allowDifferentIndexing) Then
                compareSizes = False
                Exit For
            End If
        Next dimension
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr1", arr1), KeyValue("arr2", arr2), _
                                                KeyValue("allowDifferentIndexing", allowDifferentIndexing))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function

Private Function compareDimensions(arr1 As Variant, arr2 As Variant) As Boolean
    compareDimensions = (countDimensions(arr1) = countDimensions(arr2))
End Function

Private Function compareDimensionsSize(arr1 As Variant, arr2 As Variant, _
                                        dimension As Integer, allowDifferentIndexing As Boolean) As Boolean
    Const METHOD_NAME As String = "compareDimensionsSize"
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler

    If getSize(arr1, dimension) = getSize(arr2, dimension) Then
        If allowDifferentIndexing Or (LBound(arr1, dimension) = LBound(arr2, dimension)) Then
            compareDimensionsSize = True
        End If
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr1", arr1), KeyValue("arr2", arr2), _
                                                KeyValue("allowDifferentIndexing", allowDifferentIndexing))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns 2D array containing low and up bounds for each dimension of the
'                           given parameter [arr].
' Returned value:       Array containing low and up bound for each dimension of the given array [arr].
'                       Examples:
'                           * for not initialized dynamic array function returns empty array,
'                           * for 1D array (0 To 3), function returns 2D array with single row [0, 3],
'                           * for 2D array (1 To 2, 1 To 5), function returns 2D array with two rows:
'                             [1, 2] and [1, 5]
' Parameters:
'   * arr               Array which bounds are to be obtained.
' Exceptions:
'   * NotArrayException             Thrown if the given parameter [arr] is not an array.
    '----------------------------------------------------------------------------------------------------------
Public Function getDimensionsBounds(arr As Variant) As Long()
    Const METHOD_NAME As String = "getDimensionsBounds"
    '------------------------------------------------------------------------------------------------------
    Dim result() As Long
    Dim sourceArrayDimensions As Integer
    Dim dimension As Integer
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(arr) Then Call Exceptions.NotArray.raise

    If isDefined(arr) Then
        sourceArrayDimensions = countDimensions(arr)
        ReDim result(1 To sourceArrayDimensions, 1 To 2)
        For dimension = 1 To sourceArrayDimensions
            result(dimension, 1) = LBound(arr, dimension)
            result(dimension, 2) = UBound(arr, dimension)
        Next dimension
    End If
    
    getDimensionsBounds = result


'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function checks if the given index is not ouf of bound for the given dimension in
'                           the given array.
' Returned value:       Function returns True if the given parameter [arr] is an array, it has at least as
'                           many dimensions as the value of parameter [dimension] and the given paramter
'                           [index] is between low bound and up bound of this dimension.
'                       Function returns False if:
'                           * given parameter [arr] is not an array or is dynamic array that has not been
'                             initialized yet,
'                           * given parameret [arr] is an array but it has less dimensions than value of
'                             parameter [dimension],
'                           * given parameter [index] is less than low bound or greater than upper bound
'                             for the given dimension in the given array.
' Parameters:
'   * index             Index checked.
'   * dimension         Dimension checked.
'   * arr               Array for which index is checked.
    '----------------------------------------------------------------------------------------------------------
Public Function isInBounds(index As Long, dimension As Long, arr As Variant) As Boolean
    Const METHOD_NAME As String = "isInBounds"
    '------------------------------------------------------------------------------------------------------
    Dim lowBound As Long
    Dim upBound As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If isDefined(arr) And dimension > 0 Then
        If countDimensions(arr) >= dimension Then
            lowBound = LBound(arr, dimension)
            upBound = UBound(arr, dimension)
            If index >= lowBound And index <= upBound Then
                isInBounds = True
            End If
        Else
            isInBounds = False
        End If
    Else
        isInBounds = False
    End If
       
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("index", index), _
                                                    KeyValue("dimension", dimension), KeyValue("arr", arr))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function







'[MANAGING DATA ITEMS]

'----------------------------------------------------------------------------------------------------------
' Description:          Method to add given value to the given array [arr].
'                       Parameter [value] must be compatible with the data type of given array, i.e. value
'                       of Object type cannot be added to array declared as String().
'                       * If the given array is dynamic array not initialized yet, it will initialized
'                         as (1 To 1) array with the given [value] inserted into arr(1).
'                       * If the given array is dynamic 1D array its upper bound is expanded by one and the
'                         given [value] is inserted into this slot.
'                       * In any other scenario different exceptions are thrown (details can be found in
'                         Exceptions section)
'
' Parameters:
'   * arr               Array to which [value] is being appended.
'   * value             Value appended to the given array [arr].
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array nor Empty.
'   * ExpectedDynamicArrayException     Thrown if the given parameter [arr] is not a dynamic array,
'                                       therefore adding items is not allowed.
'   * WrongDimensionsException          Thrown if the given array has more than one dimension.
'   * IllegalDataTypeException          Thrown if the given value is of data type that cannot be inserted
'                                       to the given array [arr], i.e. array is declared as Long() and
'                                       value is object.
'----------------------------------------------------------------------------------------------------------
Public Sub addEntry(ByRef arr As Variant, ByVal value As Variant)
    Const METHOD_NAME As String = "addEntry"
    '------------------------------------------------------------------------------------------------------
    Const START_INDEX As Long = 1
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If f.Variables.isUndefined(arr) And Not VBA.IsObject(arr) Then
        'Dynamic array with one item will be assigned to the variable [arr] later in this method.
    ElseIf f.Variables.isPrimitive(arr) Then
        Call Exceptions.NotArray.raise
    ElseIf VBA.IsObject(arr) Then
        Call Exceptions.NotArray.raise
    ElseIf Not isDynamic(arr) Then
        Call Exceptions.DynamicArrayExpected.raise
    ElseIf countDimensions(arr) > 1 Then
        Call Exceptions.InvalidDimension.raise
    End If

    If isDefined(arr) Then
        ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
    Else
        ReDim arr(START_INDEX To START_INDEX)
    End If
    
    On Error GoTo AssigningExceptionHandler
    If VBA.IsObject(value) Then
        Set arr(UBound(arr)) = value
    Else
        arr(UBound(arr)) = value
    End If

'==========================================================================================================
ExitPoint:
    Exit Sub
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), KeyValue("value", value))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)

    '----------------------------------------------------------------------------------------------------------
AssigningExceptionHandler:
    '[Restore previous array size] TODO
    errNumber = Exceptions.IllegalDataType.getNumber
    errDescription = Exceptions.IllegalDataType.getDescription
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), KeyValue("value", value))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Sub

Private Function getEmptyDynamicVariantArray() As Variant()
    Dim arr() As Variant
    '------------------------------------------------------------------------------------------------------
    ReDim arr(1 To 1)
    getEmptyDynamicVariantArray = arr
End Function





'[COMPARING ARRAYS]

    '----------------------------------------------------------------------------------------------------------
' Description:          Function compares two given arrays and returns True if they have the same size and
'                       the same values in each slot.
'
' Returned values:      True - if both parameters [arr1] and [arr2] are arrays with the same dimensions and
'                           size (and indexing unless paraemter [ignoreIndexingBase] is set to True) and
'                           they have equal values in each single slot.
'                       False in following cases:
'                           * at least one of the parameters [arr1], [arr2] is not an array,
'                           * parameters [arr1] and [arr2] are both arrays but they have different number
'                             of dimensions, different size in any dimension, or different indexing in
'                             any dimension (if parameter [ignoreInde xingBase] is set to False),
'                           * arrays [arr1] and [arr2] have different value in any slot; if parameter
'                             [strictDataType] is set to True, values must be of exactly the same type,
'                             otherwise function consider "2" and 2 to be equal.
'
' Parameters:
'   * arr1, arr2        Arrays to be compared.
'   * ignoreIndexingBase
'                       (Optional) If this parameter is set to True, indexing of each dimension can differ
'                       between compared arrays as long as they have the same size in this dimension.
'   * strictDataType    (Optional) If this parameter is set to True, values in each slot must have the same
'                       data type in both arrays.
    '----------------------------------------------------------------------------------------------------------
Public Function compare(arr1 As Variant, arr2 As Variant, _
                                    Optional ignoreIndexingBase As Boolean = True, _
                                    Optional strictDataType As Boolean = False) As Boolean
    Const METHOD_NAME As String = "compare"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(arr1) Then
        compare = False
    ElseIf Not VBA.IsArray(arr2) Then
        compare = False
    Else
        If compareSizes(arr1, arr2, ignoreIndexingBase) Then
            Select Case countDimensions(arr1)
                Case 0:         compare = True
                Case 1:         compare = compare1DArrays(arr1, arr2, strictDataType)
                Case 2:         compare = compare2DArrays(arr1, arr2, strictDataType)
                Case Else:      compare = False
            End Select
        End If
    End If

'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr1", arr1), KeyValue("arr2", arr2), _
            KeyValue("ignoreIndexingBase", ignoreIndexingBase), KeyValue("strictDataType", strictDataType))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function

Private Function compare1DArrays(arr1 As Variant, arr2 As Variant, strictDataType As Boolean) As Boolean
    Const METHOD_NAME As String = "compare1DArrays"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim baseDifference As Long:         baseDifference = LBound(arr2) - LBound(arr1)
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    compare1DArrays = True
    For i = LBound(arr1) To UBound(arr1)
        If Not f.Variables.equals(arr1(i), arr2(i + baseDifference), strictDataType) Then
            compare1DArrays = False
            Exit For
        End If
    Next i

'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr1", arr1), KeyValue("arr2", arr2), _
                                                                KeyValue("strictDataType", strictDataType))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function

Private Function compare2DArrays(arr1 As Variant, arr2 As Variant, strictDataType As Boolean) As Boolean
    Const METHOD_NAME As String = "compare2DArrays"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    Dim value1 As Variant
    Dim value2 As Variant
    Dim baseDifferences() As Long:      baseDifferences = getBaseDifferencesFor2DArray(arr1, arr2)
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    compare2DArrays = True
    For i = LBound(arr1, 1) To UBound(arr1, 1)
        For j = LBound(arr1, 2) To UBound(arr1, 2)
            Call f.Variables.assign(value1, arr1(i, j))
            Call f.Variables.assign(value2, arr2(i + baseDifferences(1), j + baseDifferences(2)))
            If Not f.Variables.equals(value1, value2, strictDataType) Then
                compare2DArrays = False
                Exit For
            End If
        Next j
    Next i

'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr1", arr1), KeyValue("arr2", arr2), _
                                                                KeyValue("strictDataType", strictDataType))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function

Private Function getBaseDifferencesFor2DArray(arr1 As Variant, arr2 As Variant) As Long()
    Const METHOD_NAME As String = "getBaseDifferencesFor2DArray"
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    Dim arr(1 To 2) As Long
    arr(1) = LBound(arr2, 1) - LBound(arr1, 1)
    arr(2) = LBound(arr2, 2) - LBound(arr1, 2)
    getBaseDifferencesFor2DArray = arr
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr1", arr1), KeyValue("arr2", arr2))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function






'[CONVERTING ARRAYS TO OTHER CONTAINERS]

    '----------------------------------------------------------------------------------------------------------
' Description:          Function converts the given array to string.
' Returned value:       String representation of the given array [arr].
'
' Parameters:
'   * arr               Array to be converted to string.
'   * stringifyMode     (Optional) Determines conversion mode. There are 3 available modes:
'                           - StringifyMode_Normal: no specific modification are made,
'                           - StringifyMode_Xml: array is converted to XML string, where each row is a
'                               single node and each column a subnode,
'                           - StringifyMode_Db: each item is converted to format required by database
'                               (strings are braced in apostrophes etc.)
'   * columnSeparator   (Optional) String used to separate array columns in the returned text.
'   * columnStartTag    (Optional) String used to mark the start of each column in the returned text.
'   * columnEndTag      (Optional) String used to mark the end of each column in the returned text.
'   * rowSeparator      (Optional) String used to separate array rows in the returned text.
'   * rowStartTag       (Optional) String used to mark the start of each row in the returned text.
'   * rowEndTag         (Optional) String used to mark the end of each row in the returned text.
'   * skipHeader        (Optional) Determines if the top row should be skipped when generating text
'                           representation of the given array.
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
    '----------------------------------------------------------------------------------------------------------
Public Function toString(arr As Variant, _
                                Optional stringifyMode As StringifyModeEnum = StringifyMode_Normal, _
                                Optional columnSeparator As String = ", ", _
                                Optional columnStartTag As String = vbNullString, _
                                Optional columnEndTag As String = vbNullString, _
                                Optional rowSeparator As String = vbCrLf, _
                                Optional rowStartTag As String = "[", _
                                Optional rowEndTag As String = "]", _
                                Optional skipHeader As Boolean = False) As String
    Const METHOD_NAME As String = "toString"
    '------------------------------------------------------------------------------------------------------
    Const MULTIDIMENSIONAL_ARRAY_TAG As String = "[{0}-dimensional array]"
    '------------------------------------------------------------------------------------------------------
    Dim dimensions As Integer:
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(arr) Then Call Exceptions.NotArray.raise

    dimensions = countDimensions(arr)
    Select Case dimensions
        Case 0:         toString = vbNullString
        Case 1:         toString = arrayToString_1D(arr, stringifyMode, columnSeparator, columnStartTag, _
                                                            columnEndTag, skipHeader)
        Case 2:         toString = arrayToString_2D(arr, stringifyMode, rowSeparator, rowStartTag, _
                                                            rowEndTag, columnSeparator, columnStartTag, _
                                                            columnEndTag, skipHeader)
        Case Is > 2:    toString = f.strings.format(MULTIDIMENSIONAL_ARRAY_TAG, dimensions)
    End Select

    If stringifyMode = StringifyMode_Xml Then
        Stop
        'toString = F.Strings.format(XML_ARRAY, toString)
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
            KeyValue("stringifyMode", stringifyMode), KeyValue("columnSeparator", columnSeparator), _
            KeyValue("columnStartTag", columnStartTag), KeyValue("columnEndTag", columnEndTag), _
            KeyValue("rowSeparator", rowSeparator), KeyValue("rowStartTag", rowStartTag), _
            KeyValue("rowEndTag", rowEndTag), KeyValue("skipHeader", skipHeader))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function

Private Function arrayToString_1D(arr As Variant, stringifyMode As StringifyModeEnum, _
                        ByVal itemSeparator As String, ByVal itemStartTag As String, _
                        ByVal itemEndTag As String, Optional skipHeader As Boolean = False) As String
    Const METHOD_NAME As String = "arrayToString_1D"
    '------------------------------------------------------------------------------------------------------
    Dim items() As String
    Dim itemText As String
    Dim i As Long
    Dim size As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If stringifyMode = StringifyMode_Xml Then
        itemSeparator = vbNullString
        itemStartTag = "<item>"
        itemEndTag = "</item>"
    End If

    size = UBound(arr) - LBound(arr) + 1
    If size - VBA.IIf(skipHeader, 1, 0) > 0 Then
        ReDim items(LBound(arr) + VBA.IIf(skipHeader, 1, 0) To UBound(arr))
        For i = LBound(arr) + VBA.IIf(skipHeader, 1, 0) To UBound(arr)
            itemText = f.strings.stringify(arr(i), stringifyMode)
            items(i) = itemStartTag & itemText & itemEndTag
        Next i
        arrayToString_1D = VBA.join(items, itemSeparator)
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
            KeyValue("stringifyMode", stringifyMode), KeyValue("itemSeparator", itemSeparator), _
            KeyValue("itemStartTag", itemStartTag), KeyValue("itemEndTag", itemEndTag), _
            KeyValue("skipHeader", skipHeader))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function

Private Function arrayToString_2D(arr As Variant, stringifyMode As StringifyModeEnum, _
                    ByVal rowSeparator As String, ByVal rowStartTag As String, ByVal rowEndTag As String, _
                    ByVal colSeparator As String, ByVal colStartTag As String, ByVal colEndTag As String, _
                    Optional skipHeader As Boolean = False) As String
    Const METHOD_NAME As String = "arrayToString_2D"
    '------------------------------------------------------------------------------------------------------
    Dim Rows() As String
    Dim row As Long
    Dim size As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If stringifyMode = StringifyMode_Xml Then
        rowSeparator = vbNullString
        rowStartTag = "<row>"
        rowEndTag = "</row>"
        colSeparator = vbNullString
        colStartTag = "<item>"
        colEndTag = "</item>"
    End If

    size = UBound(arr, 1) - LBound(arr, 1) + 1
    If size - VBA.IIf(skipHeader, 1, 0) > 0 Then
        ReDim Rows(LBound(arr, 1) + VBA.IIf(skipHeader, 1, 0) To UBound(arr, 1))
        For row = LBound(arr, 1) + VBA.IIf(skipHeader, 1, 0) To UBound(arr, 1)
            Rows(row) = rowStartTag & get2DArrayRowAsString(arr, stringifyMode, row, _
                                                        colSeparator, colStartTag, colEndTag) & rowEndTag
        Next row
        arrayToString_2D = VBA.join(Rows, rowSeparator)
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
            KeyValue("stringifyMode", stringifyMode), KeyValue("colSeparator", colSeparator), _
            KeyValue("colStartTag", colStartTag), KeyValue("colEndTag", colEndTag), _
            KeyValue("rowSeparator", rowSeparator), KeyValue("rowStartTag", rowStartTag), _
            KeyValue("rowEndTag", rowEndTag), KeyValue("skipHeader", skipHeader))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function

Private Function get2DArrayRowAsString(arr As Variant, stringifyMode As StringifyModeEnum, row As Long, _
                        itemsDelimiter As String, itemStartTag As String, itemEndTag As String) As String
    Const METHOD_NAME As String = "get2DArrayRowAsString"
    '------------------------------------------------------------------------------------------------------
    Dim items() As String
    Dim column As Long
    Dim itemText As String
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    ReDim items(LBound(arr, 2) To UBound(arr, 2))
    For column = LBound(arr, 2) To UBound(arr, 2)
        itemText = f.strings.stringify(arr(row, column), stringifyMode)
        items(column) = itemStartTag & itemText & itemEndTag
    Next column
    get2DArrayRowAsString = VBA.join(items, itemsDelimiter)
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
            KeyValue("stringifyMode", stringifyMode), KeyValue("itemsDelimiter", itemsDelimiter), _
            KeyValue("itemStartTag", itemStartTag), KeyValue("itemEndTag", itemEndTag))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function converts the given 1D array to collection.
' Returned value:       Collection containing the same values as the given array [arr].
'
' Parameters:
'   * arr               Array to be converted to collection.
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
'   * WrongDimensionsNumberException    Thrown if the given array has more than one dimension.
    '----------------------------------------------------------------------------------------------------------
Public Function toCollection(arr As Variant) As VBA.Collection
    Const METHOD_NAME As String = "toCollection"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(arr) Then Call Exceptions.NotArray.raise
    If f.arrays.countDimensions(arr) > 1 Then Call Exceptions.InvalidDimension.raise

    Set toCollection = New VBA.Collection
    If isDefined(arr) Then
        For i = LBound(arr) To UBound(arr)
            Call toCollection.Add(arr(i))
        Next i
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function




    '----------------------------------------------------------------------------------------------------------
' Description:          Function converts the given value to 2D array.
' Returned value:       Depending on the input parameter [value] function returns the following result:
'                        * 1D array - function returns 2D array with a single column and as many rows as
'                               the number of items in the original 1D array,
'                        * 2D array - function returns clone of the original array,
'                        * array with over 2 dimensions - function throws WrongDimensionsNumberException
'                        * dynamic array not initialized - function throws WrongDimensionsNumberException
'                        * dictionary - function returns 2D array with two columns (one column for key,
'                               the second one for value).
'                        * collection - function returns 2D array with a single column and as many rows as
'                               the number of items in the given collection,
'                        * primitive value or Nothing - function returns 2D array with one row and one
'                               column and original value inserted in this single slot.
'
' Parameters:
'   * arr               Array to be converted to collection.
'
' Exceptions:
'   * WrongDimensionsNumberException    Thrown if the given value is an array with more than two dimensions
'                                       or if it is dynamic array that has not been initialized yet.
    '----------------------------------------------------------------------------------------------------------
Public Function to2D(value As Variant) As Variant
    Const METHOD_NAME As String = "to2D"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If VBA.IsArray(value) Then
        Select Case countDimensions(value)
            Case 1:             to2D = convertTo2DArray_From1DArray(value)
            Case 0, 2:          to2D = value
            Case Else:          Call Exceptions.InvalidDimension.raise
        End Select
    ElseIf VBA.IsObject(value) Then
        If value Is Nothing Then
            to2D = convertTo2DArray_FromSingleValue(value)
        ElseIf TypeOf value Is Scripting.Dictionary Then
            to2D = convertTo2DArray_FromDictionary(value)
        ElseIf TypeOf value Is VBA.Collection Then
            to2D = convertTo2DArray_FromCollection(value)
        Else
            to2D = convertTo2DArray_FromSingleValue(value)
        End If
    Else
        to2D = convertTo2DArray_FromSingleValue(value)
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("value", value))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function

Public Function to2DArray(value As Variant) As Variant
    to2DArray = to2D(value)
End Function

Private Function convertTo2DArray_From1DArray(sourceArr As Variant) As Variant
    Const METHOD_NAME As String = "convertTo2DArray_From1DArray"
    '------------------------------------------------------------------------------------------------------
    Dim arr As Variant
    Dim lngRow As Long
    Dim lngLBound As Long
    Dim lngUBound As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    ReDim arr(1 To UBound(sourceArr) - LBound(sourceArr) + 1, 1 To 1)
    lngLBound = LBound(sourceArr)
    lngUBound = UBound(sourceArr)

    For lngRow = lngLBound To lngUBound
        Call f.Variables.assign(arr(lngRow - lngLBound + 1, 1), sourceArr(lngRow))
    Next lngRow
    convertTo2DArray_From1DArray = arr


'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("sourceArr", sourceArr))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)

End Function

Private Function convertTo2DArray_FromDictionary(value As Variant) As Variant
    Const METHOD_NAME As String = "convertTo2DArray_FromDictionary"
    '------------------------------------------------------------------------------------------------------
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    Dim arr() As Variant
    Dim Counter As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If f.dictionaries.isDictionary(value) Then
        Set dict = value
    End If

    If f.dictionaries.isNonEmptyDictionary(dict) Then
        ReDim arr(1 To dict.count, 1 To 2)
        For Each key In dict.Keys
            Counter = Counter + 1
            Call f.Variables.assign(arr(Counter, 1), key)
            Call f.Variables.assign(arr(Counter, 2), dict.item(key))
        Next key
    End If

    convertTo2DArray_FromDictionary = arr

'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("value", value))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)

End Function

Private Function convertTo2DArray_FromCollection(value As Variant) As Variant
    Const METHOD_NAME As String = "convertTo2DArray_FromCollection"
    '------------------------------------------------------------------------------------------------------
    Dim col As Object
    Dim item As Variant
    Dim arr() As Variant
    Dim Counter As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler

    If f.Collections.isCollection(value) Then
        Set col = value
    End If

    If f.Collections.isNonEmptyCollection(col) Then
        ReDim arr(1 To col.count, 1 To 1)
        For Each item In col
            Counter = Counter + 1
            Call f.Variables.assign(arr(Counter, 1), item)
        Next item
    End If

    convertTo2DArray_FromCollection = arr


'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("value", value))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)

End Function

Private Function convertTo2DArray_FromSingleValue(value As Variant) As Variant
    Const METHOD_NAME As String = "convertTo2DArray_FromSingleValue"
    '------------------------------------------------------------------------------------------------------
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    ReDim arr(1 To 1, 1 To 1)
    Call f.Variables.assign(arr(1, 1), value)
    convertTo2DArray_FromSingleValue = arr


'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("value", value))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)

End Function






'[TESTING FOR ITEMS]
    '----------------------------------------------------------------------------------------------------------
' Description:          Function checks if the given 1D arr contains the given value.
' Returned value:       True - if the given array [arr] contains the given value [value]. If optional
'                           parameter [strictDataType] is set to True, given value and value found in array
'                           must have exactly the same data type, otherwise they must be equal but can have
'                           other data type (i.e. "2" and 2).
'
' Parameters:
'   * arr               Array to be searched.
'   * value             Value that function looks for in the given array.
'   * strictDataType    (Optional) Determines if the value in array and searched value must be of exactly
'                       the same data type.
'   * isCaseSensitive   (Optional) Determines if letter-casing is ignored when searching through the array.
'   * trimIfValueIsString
'                       (Optional) If given parameter [value] is string, this parameter determines if
'                       blank spaces should be ignored when searching array.
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
'   * WrongDimensionsNumberException    Thrown if the given array has more than 1 dimension or if it is
'                                       dynamic array that has not been initialized yet.
'   * IllegalDataTypeException          Thrown if the parameter [value] is an array. It must be object or
'                                       primitive type.
    '----------------------------------------------------------------------------------------------------------
Public Function contains(arr As Variant, value As Variant, _
                                Optional strictDataType As Boolean = True, _
                                Optional isCaseSensitive As Boolean = False, _
                                Optional trimIfValueIsString As Boolean = True) As Boolean
    Const METHOD_NAME As String = "contains"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(arr) Then Call Exceptions.NotArray.raise
    If f.arrays.countDimensions(arr) > 1 Then Call Exceptions.InvalidDimension.raise
    If VBA.IsArray(value) Then Call Exceptions.IllegalValue.raiseWithCustomMessage("Value cannot be of array type")

    If isDefined(arr) Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If f.Variables.equals(arr(i), value, strictDataType, isCaseSensitive, trimIfValueIsString) Then
                contains = True
                Exit For
            End If
        Next i
    End If

'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("value", value), _
                KeyValue("strictDataType", strictDataType), KeyValue("isCaseSensitive", isCaseSensitive), _
                KeyValue("trimIfValueIsString", trimIfValueIsString))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)

End Function







'[SLICING ARRAYS]
    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns the first item stored in the given array [arr].
' Returned value:       First item stored in the given array (item having index for each dimension equal to
'                           LBound of this dimension, unless it is undefined/emptyString/zero and
'                           parameters [skipUndefinedItems], [skipEmptyStrings] and [skipZeros] excludes it
'                           from being returned - in such case the next item that meets the requirements
'                           is returned). If there is no item that meets requirements, Empty is returned.
'
' Parameters:
'   * arr               Array to be searched.
'   * skipUndefinedItems
'                       (Optional) If this parameter is set to True each item with undefined value
'                       (Empty/Missing/Null/Nothing) is skipped and the first not undefined value is
'                       returned.
'   * skipEmptyStrings  (Optional) As above but for empty strings.
'   * skipZeros         (Optional) As above but for zeroes.
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
'   * WrongDimensionsNumberException    Thrown if the given array has more than 2 dimension.
'   * UndefinedArrayNotAllowedException Thrown if the parameter [arr] is a dynamic array that has not been
'                                       initialized yet.
    '----------------------------------------------------------------------------------------------------------
Public Function getFirstItem(arr As Variant, _
                                            Optional skipUndefinedItems As Boolean = True, _
                                            Optional skipEmptyStrings As Boolean = True, _
                                            Optional skipZeros As Boolean = True) As Variant
    Const METHOD_NAME As String = "getFirstItem"
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If Not VBA.IsArray(arr) Then Call Exceptions.NotArray.raise
    If Not isDefined(arr) Then Call Exceptions.UndefinedArrayNotAllowed.raise
    If f.arrays.countDimensions(arr) > 2 Then Call Exceptions.InvalidDimension.raise
    Select Case countDimensions(arr)
        Case 1: Call f.Variables.assign(getFirstItem, _
                                getFirstItem_1DArray(arr, skipUndefinedItems, skipEmptyStrings, skipZeros))
        Case 2: Call f.Variables.assign(getFirstItem, _
                                getFirstItem_2DArray(arr, skipUndefinedItems, skipEmptyStrings, skipZeros))
    End Select
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
                KeyValue("skipUndefinedItems", skipUndefinedItems), _
                KeyValue("skipEmptyStrings", skipEmptyStrings), KeyValue("skipZeros", skipZeros))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function

Private Function getFirstItem_1DArray(arr As Variant, skipUndefinedItems As Boolean, _
                                            skipEmptyStrings As Boolean, skipZeros As Boolean) As Variant
    Const METHOD_NAME As String = "getFirstItem_1DArray"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    For i = LBound(arr) To UBound(arr)
        If getItem_PassTest(arr(i), skipUndefinedItems, skipEmptyStrings, skipZeros) Then
            Call f.Variables.assign(getFirstItem_1DArray, arr(i))
            Exit For
        End If
    Next i
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
                KeyValue("skipUndefinedItems", skipUndefinedItems), _
                KeyValue("skipEmptyStrings", skipEmptyStrings), KeyValue("skipZeros", skipZeros))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)

End Function

Private Function getFirstItem_2DArray(arr As Variant, skipUndefinedItems As Boolean, _
                                            skipEmptyStrings As Boolean, skipZeros As Boolean) As Variant
    Const METHOD_NAME As String = "getFirstItem_2DArray"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    Dim isFound As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If getItem_PassTest(arr(i, j), skipUndefinedItems, skipEmptyStrings, skipZeros) Then
                Call f.Variables.assign(getFirstItem_2DArray, arr(i, j))
                isFound = True
                Exit For
            End If
        Next j
        If isFound Then Exit For
    Next i
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
                KeyValue("skipUndefinedItems", skipUndefinedItems), _
                KeyValue("skipEmptyStrings", skipEmptyStrings), KeyValue("skipZeros", skipZeros))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
    
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns the last item stored in the given array [arr].
' Returned value:       Last item stored in the given array (item having index for each dimension equal to
'                           LBound of this dimension, unless it is undefined/emptyString/zero and
'                           parameters [skipUndefinedItems], [skipEmptyStrings] and [skipZeros] excludes it
'                           from being returned - in such case the previous item that meets the requirements
'                           is returned). If there is no item that meets requirements, Empty is returned.
'
' Parameters:
'   * arr               Array to be searched.
'   * skipUndefinedItems
'                       (Optional) If this parameter is set to True each item with undefined value
'                       (Empty/Missing/Null/Nothing) is skipped and the previous not undefined value is
'                       returned.
'   * skipEmptyStrings  (Optional) As above but for empty strings.
'   * skipZeros         (Optional) As above but for zeroes.
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
'   * WrongDimensionsNumberException    Thrown if the given array has more than 2 dimension.
'   * UndefinedArrayNotAllowedException Thrown if the parameter [arr] is a dynamic array that has not been
'                                       initialized yet.
    '----------------------------------------------------------------------------------------------------------
Public Function getLastItem(arr As Variant, _
                                            Optional skipUndefinedItems As Boolean = True, _
                                            Optional skipEmptyStrings As Boolean = True, _
                                            Optional skipZeros As Boolean = True) As Variant
    Const METHOD_NAME As String = "getLastItem"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(arr) Then Call Exceptions.NotArray.raise
    If Not isDefined(arr) Then Call Exceptions.UndefinedArrayNotAllowed.raise
    If f.arrays.countDimensions(arr) > 2 Then Call Exceptions.InvalidDimension.raise
    Select Case countDimensions(arr)
        Case 1: Call f.Variables.assign(getLastItem, _
                                getLastItem_1DArray(arr, skipUndefinedItems, skipEmptyStrings, skipZeros))
        Case 2: Call f.Variables.assign(getLastItem, _
                                getLastItem_2DArray(arr, skipUndefinedItems, skipEmptyStrings, skipZeros))
    End Select
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
                KeyValue("skipUndefinedItems", skipUndefinedItems), _
                KeyValue("skipEmptyStrings", skipEmptyStrings), KeyValue("skipZeros", skipZeros))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function

Private Function getLastItem_1DArray(arr As Variant, skipUndefinedItems As Boolean, _
                                            skipEmptyStrings As Boolean, skipZeros As Boolean) As Variant
    Const METHOD_NAME As String = "getLastItem_1DArray"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    For i = UBound(arr) To LBound(arr) Step -1
        If getItem_PassTest(arr(i), skipUndefinedItems, skipEmptyStrings, skipZeros) Then
            Call f.Variables.assign(getLastItem_1DArray, arr(i))
            Exit For
        End If
    Next i
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
                KeyValue("skipUndefinedItems", skipUndefinedItems), _
                KeyValue("skipEmptyStrings", skipEmptyStrings), KeyValue("skipZeros", skipZeros))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function

Private Function getLastItem_2DArray(arr As Variant, skipUndefinedItems As Boolean, _
                                            skipEmptyStrings As Boolean, skipZeros As Boolean) As Variant
    Const METHOD_NAME As String = "getLastItem_2DArray"
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    Dim isFound As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    For i = UBound(arr, 1) To LBound(arr, 1) Step -1
        For j = UBound(arr, 2) To LBound(arr, 2) Step -1
            If getItem_PassTest(arr(i, j), skipUndefinedItems, skipEmptyStrings, skipZeros) Then
                Call f.Variables.assign(getLastItem_2DArray, arr(i, j))
                isFound = True
                Exit For
            End If
        Next j
        If isFound Then Exit For
    Next i
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
                KeyValue("skipUndefinedItems", skipUndefinedItems), _
                KeyValue("skipEmptyStrings", skipEmptyStrings), KeyValue("skipZeros", skipZeros))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function

Private Function getItem_PassTest(item As Variant, skipUndefinedItems As Boolean, _
                                            skipEmptyStrings As Boolean, skipZeros As Boolean) As Boolean
    Const METHOD_NAME As String = "getItem_PassTest"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If f.Variables.isUndefined(item) Then
        getItem_PassTest = (Not skipUndefinedItems)
    ElseIf VBA.IsObject(item) Then
        getItem_PassTest = True
    ElseIf VBA.IsArray(item) Then
        getItem_PassTest = True
    ElseIf item = vbNullString Then
        getItem_PassTest = (Not skipEmptyStrings)
    ElseIf item = 0 Then
        getItem_PassTest = (Not skipZeros)
    Else
        getItem_PassTest = True
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("item", item), _
                KeyValue("skipUndefinedItems", skipUndefinedItems), _
                KeyValue("skipEmptyStrings", skipEmptyStrings), KeyValue("skipZeros", skipZeros))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns a part of the given 1D array beginning from an item defined by
'                       [startIndex] and ending to the item defined by [endIndex] (including this item).
' Returned value:       1D array containing all items from the original array [arr] between [startIndex]
'                       and [endIndex].
'                       If parameter [absoluteIndex] is set to True, items in the returned array has the
'                       same indexing as in the original array, i.e.:
'                       if original array is String(1 To 10), [startIndex] = 3, [endIndex] = 6 the returned
'                       array will be String(3 To 6).
'                       If parameter [absoluteIndex] is set to False, the returned array is indexed from 1.
'
'                       If startIndex is greater than endIndex items are returned in reversed order.
'
' Parameters:
'   * arr               Original array from which subarray is retrieved.
'   * startIndex, endIndex
'                       Indices limiting the range of the returned array.
'   * absoluteIndex     (Optional) If this value is set to True, items in the returned array has the
'                       same indexing as in the original array. Otherwise they are indexed from 1.
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
'   * UndefinedArrayNotAllowedException Thrown if the parameter [arr] is a dynamic array that has not been
'                                       initialized yet.
'   * WrongDimensionsNumberException    Thrown if the given array has more than 1 dimension.
    '----------------------------------------------------------------------------------------------------------
Public Function slice1D_ByEndIndex(arr As Variant, startIndex As Long, endIndex As Long, _
                                 Optional absoluteIndex As Boolean = True) As Variant
    Const METHOD_NAME As String = "slice1D_ByEndIndex"
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(arr) Then Call Exceptions.NotArray.raise
    If Not f.arrays.isDefined(arr) Then Call Exceptions.UndefinedArrayNotAllowed.raise
    If f.arrays.countDimensions(arr) > 1 Then Call Exceptions.InvalidDimension.raise

    If absoluteIndex Then
        slice1D_ByEndIndex = slice1D_ByEndIndex_absoluteIndex(arr, startIndex, endIndex)
    Else
        slice1D_ByEndIndex = slice1D_ByEndIndex_relativeIndex(arr, startIndex, endIndex)
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
                KeyValue("startIndex", startIndex), KeyValue("endIndex", endIndex), _
                KeyValue("absoluteIndex", absoluteIndex))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function

Private Function slice1D_ByEndIndex_absoluteIndex(sourceArray As Variant, _
                                                        startIndex As Long, endIndex As Long) As Variant
    Const METHOD_NAME As String = "slice1D_ByEndIndex_absoluteIndex"
    '------------------------------------------------------------------------------------------------------
    Dim arr As Variant
    Dim i As Long
    Dim isReversed As Boolean
    Dim lowerIndex As Long
    Dim upperIndex As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not isInBounds(startIndex, 1, sourceArray) Then Call Exceptions.IndexOutOfBound.raiseWithCustomMessage("Parameter [startIndex] is out of range")
    If Not isInBounds(endIndex, 1, sourceArray) Then Call Exceptions.IndexOutOfBound.raiseWithCustomMessage("Parameter [endIndex] is out of range")

    isReversed = startIndex > endIndex
    If isReversed Then
        lowerIndex = endIndex
        upperIndex = startIndex
        ReDim arr(endIndex To startIndex)
    Else
        lowerIndex = startIndex
        upperIndex = endIndex
        ReDim arr(startIndex To endIndex)
    End If

    For i = lowerIndex To upperIndex
        If isReversed Then
            Call f.Variables.assign(arr(i), sourceArray(upperIndex - i + lowerIndex))
        Else
            Call f.Variables.assign(arr(i), sourceArray(i))
        End If
    Next i

    slice1D_ByEndIndex_absoluteIndex = arr



'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("sourceArray", sourceArray), _
                KeyValue("startIndex", startIndex), KeyValue("endIndex", endIndex))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function

Private Function slice1D_ByEndIndex_relativeIndex(sourceArray As Variant, _
                                                            startIndex As Long, endIndex As Long) As Variant
    Const METHOD_NAME As String = "slice1D_ByEndIndex_relativeIndex"
    '------------------------------------------------------------------------------------------------------
    Dim sourceSize As Long
    Dim destinationSize As Long
    Dim baseDifference As Long
    Dim isReversed As Boolean
    Dim sourceArrayIndex As Long
    Dim arr As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    sourceSize = getSize(sourceArray, 1)
    If Not f.Numerics.isInRange(startIndex, 1, sourceSize) Then Call Exceptions.IndexOutOfBound.raiseWithCustomMessage("Parameter [startIndex] is out of range")
    If Not f.Numerics.isInRange(endIndex, 1, sourceSize) Then Call Exceptions.IndexOutOfBound.raiseWithCustomMessage("Parameter [endIndex] is out of range")

    isReversed = startIndex > endIndex
    destinationSize = VBA.IIf(isReversed, startIndex - endIndex + 1, endIndex - startIndex + 1)
    baseDifference = 1 - LBound(sourceArray)
    ReDim arr(1 To destinationSize)
    For i = 1 To destinationSize
        sourceArrayIndex = startIndex - baseDifference + (i - 1) * VBA.IIf(isReversed, -1, 1)
        Call f.Variables.assign(arr(i), sourceArray(sourceArrayIndex))
    Next i

    slice1D_ByEndIndex_relativeIndex = arr
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("sourceArray", sourceArray), _
                KeyValue("startIndex", startIndex), KeyValue("endIndex", endIndex))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns a part of the given 2D array limited by given [startRowIndex],
'                       [endRowIndex], [startColumnIndex] and [endColumnIndex] parameters (end indices are
'                       included in the result array).
' Returned value:       2D array containing all rows from the original array [arr] between [startRowIndex]
'                       and [endRowIndex] and all columns from [startColumnIndex] and [endColumnIndex].
'
'                       If parameter [absoluteIndex] is set to True, items in the returned array has the
'                       same indexing as in the original array, i.e.:
'                       if original array is Long(1 To 8, 1 To 8), [startRowIndex] = 3, [endRowIndex] = 6,
'                       [startColumnIndex] = 2, [endColumnIndex] = 5, then the returned array will be
'                       Long(3 To 6, 2 To 5).
'                       If parameter [absoluteIndex] is set to False, each dimension in the returned array
'                       is indexed from 1.
'
'                       If startRowIndex is greater than endRowIndex then rows are returned in reversed
'                       order.
'                       If startColumnIndex is greater than endColumnIndex then columns rows are returned
'                       in reversed order.
'
' Parameters:
'   * arr               Original array from which subarray is retrieved.
'   * startRowIndex, endRowIndex, startColumnIndex, endColumnIndex
'                       Indices limiting the range of the returned array.
'   * absoluteIndex     (Optional) If this value is set to True, items in the returned array has the
'                       same indexing as in the original array. Otherwise they are indexed from 1.
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
'   * UndefinedArrayNotAllowedException Thrown if the parameter [arr] is a dynamic array that has not been
'                                       initialized yet.
'   * WrongDimensionsNumberException    Thrown if the given array has more or less than 2 dimensions.
    '----------------------------------------------------------------------------------------------------------
Public Function slice2D_ByEndIndex(sourceArray As Variant, _
                                    startRowIndex As Long, endRowIndex As Long, _
                                    startColumnIndex As Long, endColumnIndex As Long, _
                                    Optional absoluteIndex As Boolean = True) As Variant
    Const METHOD_NAME As String = "slice2D_ByEndIndex"
    '------------------------------------------------------------------------------------------------------
    Dim realStartRowIndex As Long
    Dim realEndRowIndex As Long
    Dim realStartColumnIndex As Long
    Dim realEndColumnIndex As Long
    Dim arr() As Variant
    Dim Rows As Long
    Dim columns As Long
    Dim i As Long
    Dim row As Long
    Dim j As Long
    Dim column As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(sourceArray) Then Call Exceptions.NotArray.raise
    If Not isDefined(sourceArray) Then Call Exceptions.UndefinedArrayNotAllowed.raise
    If countDimensions(sourceArray) <> 2 Then Exceptions.InvalidDimension.raise

    If absoluteIndex Then
        realStartRowIndex = startRowIndex
        realEndRowIndex = endRowIndex
        realStartColumnIndex = startColumnIndex
        realEndColumnIndex = endColumnIndex
    Else
        realStartRowIndex = startRowIndex + LBound(sourceArray, 1) - 1
        realEndRowIndex = endRowIndex + LBound(sourceArray, 1) - 1
        realStartColumnIndex = startColumnIndex + LBound(sourceArray, 2) - 1
        realEndColumnIndex = endColumnIndex + LBound(sourceArray, 2) - 1
    End If

    If Not isInBounds(realStartRowIndex, 1, sourceArray) Then
        Call Exceptions.IndexOutOfBound.raise
    ElseIf Not isInBounds(realEndRowIndex, 1, sourceArray) Then
        Call Exceptions.IndexOutOfBound.raise
    ElseIf Not isInBounds(realStartColumnIndex, 2, sourceArray) Then
        Call Exceptions.IndexOutOfBound.raise
    ElseIf Not isInBounds(realEndColumnIndex, 2, sourceArray) Then
        Call Exceptions.IndexOutOfBound.raise
    End If

    'Creating final array.
    Rows = VBA.Abs(realStartRowIndex - realEndRowIndex) + 1
    columns = VBA.Abs(realStartColumnIndex - realEndColumnIndex) + 1

    ReDim arr(1 To Rows, 1 To columns)
    For i = 1 To Rows
        row = VBA.IIf(realEndRowIndex < realStartRowIndex, realStartRowIndex - i + 1, _
                                                                            realStartRowIndex + i - 1)
        For j = 1 To columns
            column = VBA.IIf(realEndColumnIndex < realStartColumnIndex, _
                                           realStartColumnIndex - j + 1, realStartColumnIndex + j - 1)
            Call f.Variables.assign(arr(i, j), sourceArray(row, column))
        Next j
    Next i

    slice2D_ByEndIndex = arr
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("sourceArray", sourceArray), _
                KeyValue("startRowIndex", startRowIndex), KeyValue("endRowIndex", endRowIndex), _
                KeyValue("startColumnIndex", startColumnIndex), _
                KeyValue("endColumnIndex", endColumnIndex), KeyValue("absoluteIndex", absoluteIndex))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns a single row of the given 2D array and returns it as a 1D array.
' Returned value:       1D array containing all the items from the given row in the original array [arr].
'
'                       If parameter [absoluteIndex] is set to True, the row is retrieved by index and
'                       items in the returned array has the same indexing as in the original array, i.e.:
'                       if original array is Long(2 To 8, 3 To 7), [rowIndex] = 4, then the returned array
'                       will be Long(2 To 8) having all items from row 4 - (1,4), (2,4) etc.
'
'                       If parameter [absoluteIndex] is set to False, indexing is ignored when retrieving
'                       row and the returned array is indexed from 1, i.e.:
'                       if original array is Long(2 To 8, 3 To 7), [rowIndex] = 2, then the returned array
'                       will be Long(1 To 7) having all items from row with index 4 (because this is 2nd
'                       row in the original array).
'
' Parameters:
'   * sourceArray       Original array from which subarray is retrieved.
'   * rowIndex          Index of row to be retrieved.
'   * absoluteIndex     (Optional) If this value is set to True, row is retrieved by its index. Otherwise
'                       row is retrieved by its position in original array (i.e. for rowIndex = 1 the first
'                       row is returned no matter what its index is).
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
'   * UndefinedArrayNotAllowedException Thrown if the parameter [arr] is a dynamic array that has not been
'                                       initialized yet.
'   * WrongDimensionsNumberException    Thrown if the given array has more or less than 2 dimensions.
    '----------------------------------------------------------------------------------------------------------
Public Function getRow(sourceArray As Variant, RowIndex As Long, _
                                                    Optional absoluteIndex As Boolean = True) As Variant
    Const METHOD_NAME As String = "getRow"
    '------------------------------------------------------------------------------------------------------
    Dim arr() As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(sourceArray) Then Call Exceptions.NotArray.raise
    If Not isDefined(sourceArray) Then Call Exceptions.UndefinedArrayNotAllowed.raise
    If countDimensions(sourceArray) <> 2 Then Exceptions.InvalidDimension.raise

    If absoluteIndex Then
        If Not isInBounds(RowIndex, 1, sourceArray) Then Call Exceptions.IndexOutOfBound.raise
    Else
        If Not f.Numerics.isInRange(RowIndex, 0, getSize(sourceArray, 1), False, True) Then _
                                                                   Call Exceptions.IndexOutOfBound.raise
    End If

    If absoluteIndex Then
        ReDim arr(LBound(sourceArray, 2) To UBound(sourceArray, 2))
        For i = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            Call f.Variables.assign(arr(i), sourceArray(RowIndex, i))
        Next i
    Else
        ReDim arr(1 To getSize(sourceArray, 2))
        For i = 1 To getSize(sourceArray, 2)
            Call f.Variables.assign(arr(i), _
                        sourceArray(RowIndex + LBound(sourceArray, 1) - 1, i + LBound(sourceArray, 2) - 1))
        Next i
    End If

    getRow = arr
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("sourceArray", sourceArray), _
                        KeyValue("rowIndex", RowIndex), KeyValue("absoluteIndex", absoluteIndex))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function





    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns a single column of the given 2D array and returns it as a 1D
'                       array.
'
' Returned value:       1D array containing all the items from the given column of the original array
'                       [arr].
'
'                       If parameter [absoluteIndex] is set to True, the column is retrieved by index and
'                       items in the returned array has the same indexing as in the original array, i.e.:
'                       if original array is Long(2 To 8, 3 To 7), [columnIndex] = 4, then the returned
'                       array will be Long(3 To 7) having all items from column 4 - (4,3), (4,4) etc.
'
'                       If parameter [absoluteIndex] is set to False, indexing is ignored when retrieving
'                       column and the returned array is indexed from 1, i.e.:
'                       if original array is Long(2 To 8, 3 To 7), [columnIndex] = 2, then the returned
'                       array will be Long(1 To 5) having all items from column with index 3 (because this
'                       is 2nd column in the original array).
'
' Parameters:
'   * sourceArray       Original array from which subarray is retrieved.
'   * columnIndex       Index of column to be retrieved.
'   * absoluteIndex     (Optional) If this value is set to True, column is retrieved by its index.
'                       Otherwise column is retrieved by its position in original array (i.e. for
'                       columnIndex = 1 the first column is returned no matter what its index is).
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
'   * UndefinedArrayNotAllowedException Thrown if the parameter [arr] is a dynamic array that has not been
'                                       initialized yet.
'   * WrongDimensionsNumberException    Thrown if the given array has more or less than 2 dimensions.
    '----------------------------------------------------------------------------------------------------------
Public Function getColumn(sourceArray As Variant, columnIndex As Long, _
                                                       Optional absoluteIndex As Boolean = True) As Variant
    Const METHOD_NAME As String = "getColumn"
    '------------------------------------------------------------------------------------------------------
    Dim arr() As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If Not VBA.IsArray(sourceArray) Then Call Exceptions.NotArray.raise
    If Not isDefined(sourceArray) Then Call Exceptions.UndefinedArrayNotAllowed.raise
    If countDimensions(sourceArray) <> 2 Then Exceptions.InvalidDimension.raise

    If absoluteIndex Then
        If Not isInBounds(columnIndex, 2, sourceArray) Then Call Exceptions.IndexOutOfBound.raise
    Else
        If Not f.Numerics.isInRange(columnIndex, 0, getSize(sourceArray, 2), False, True) Then _
                                                                      Call Exceptions.IndexOutOfBound.raise
    End If

    If absoluteIndex Then
        ReDim arr(LBound(sourceArray, 1) To UBound(sourceArray, 1))
        For i = LBound(sourceArray, 1) To UBound(sourceArray, 1)
            Call f.Variables.assign(arr(i), sourceArray(i, columnIndex))
        Next i
    Else
        ReDim arr(1 To getSize(sourceArray, 1))
        For i = 1 To getSize(sourceArray, 1)
            Call f.Variables.assign(arr(i), _
                    sourceArray(i + LBound(sourceArray, 1) - 1, columnIndex + LBound(sourceArray, 2) - 1))
        Next i
    End If

    getColumn = arr
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("sourceArray", sourceArray), _
                        KeyValue("columnIndex", columnIndex), KeyValue("absoluteIndex", absoluteIndex))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)
        
End Function





'[FILTERING ARRAYS]
    '----------------------------------------------------------------------------------------------------------
' Description:          Function returns only unique values from the given array.
'                       Only 1D or dynamic-not-yet-initialized arrays are accepted.
'
' Returned value:       1D array containing unique values from the original 1D array.
'                       For a dynamic array that has not been initialized yet, also dynamic-not-initialized
'                       array is returned.
'
' Parameters:
'   * arr               Original array to be filtered.
'   * strictDataType    (Optional) It determines if equal values of different data types should be
'                       considered as unique, i.e.: if this parameter is set to True, values "2" and 2 are
'                       consideded as unique and both of them will be returned.
'   * isCaseSensitive   (Optional) It determines if letter-casing is taken into consideration when
'                       filtering original array. If this value is set to True, values "a" and "A" are
'                       considered as unique and both will be returned. Otherwise they are considered as
'                       equal and only one of them will be returned (with lower index).
'   * trimWhenComparingStrings
'                       (Optional) It determines if blank spaces should be ignored when filtering the
'                       original array, i.e.: If this parameter is set to True, values "a" and " a " are
'                       considered as equal and only one of them will be returned. Otherwise they are
'                       considered as different and both will be returned.
'
' Exceptions:
'   * NotArrayException                 Thrown if the given parameter [arr] is not an array.
'   * WrongDimensionsNumberException    Thrown if the given array has more or less than 2 dimensions.
    '----------------------------------------------------------------------------------------------------------
Public Function getUniqueValues(arr As Variant, Optional strictDataType As Boolean = False, _
                                Optional isCaseSensitive As Boolean = False, _
                                Optional trimWhenComparingStrings As Boolean = True) As Variant()
    Const METHOD_NAME As String = "getUniqueValues"
    '------------------------------------------------------------------------------------------------------
    Dim result() As Variant
    Dim dict As Scripting.Dictionary
    Dim i As Long
    Dim varKey As Variant
    '------------------------------------------------------------------------------------------------------

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not VBA.IsArray(arr) Then Call Exceptions.NotArray.raise
    If countDimensions(arr) > 1 Then Call Exceptions.InvalidDimension.raise

    If isDefined(arr) Then

        Set dict = f.dictionaries.Create(isCaseSensitive)
        For i = LBound(arr) To UBound(arr)

            If f.Variables.isUndefined(arr(i)) Then
                If strictDataType Then
                    varKey = VBA.varType(arr(i)) & "|" & f.strings.stringify(arr(i))
                Else
                    varKey = Empty
                End If
            ElseIf VBA.IsObject(arr(i)) Then
                Set varKey = arr(i)
            ElseIf strictDataType Then
                If trimWhenComparingStrings Then
                    varKey = VBA.varType(arr(i)) & "|" & f.strings.trimFull(f.strings.stringify(arr(i)))
                Else
                    varKey = VBA.varType(arr(i)) & "|" & f.strings.stringify(arr(i))
                End If
            Else
                If trimWhenComparingStrings Then
                    varKey = f.strings.trimFull(f.strings.stringify(arr(i)))
                Else
                    varKey = f.strings.stringify(arr(i))
                End If
            End If

            With dict
                If Not .exists(varKey) Then
                    Call .Add(varKey, arr(i))
                End If
            End With

        Next i

        getUniqueValues = dict.items

    Else

        getUniqueValues = result

    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function
    
    '----------------------------------------------------------------------------------------------------------
ErrHandler:
    Dim err As VBA.errObject, errNumber As Long, errDescription As String, params As Scripting.Dictionary
    Set err = VBA.err: errNumber = err.number: errDescription = err.description
    Set params = f.dictionaries.createWithItems(False, KeyValue("arr", arr), _
                        KeyValue("strictDataType", strictDataType), _
                        KeyValue("isCaseSensitive", isCaseSensitive), _
                        KeyValue("trimWhenComparingStrings", trimWhenComparingStrings))
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, params)
    Call VBA.err.raise(errNumber, Source:=CLASS_NAME & "." & METHOD_NAME, description:=errDescription)

End Function







'[SORTING ARRAYS]
Public Function sort(arr As Variant, Optional sortAscending As Boolean = True, Optional sortColumn As Long = 1, Optional hasHeader As Boolean = False) As Variant
     Dim sorter As ArraySorter
    '------------------------------------------------------------------------------------------------------
    If f.arrays.isDefined(arr) Then
        Set sorter = New ArraySorter
        sort = sorter.setArray(arr).sort(ascending:=sortAscending, sortColumn:=sortColumn, hasHeader:=hasHeader)
    Else
        sort = arr
    End If
End Function



Public Sub swapItems(arr As Variant, row1 As Long, row2 As Long, Optional ByVal dimensions As Long = 0)
    If dimensions = 0 Then dimensions = f.arrays.countDimensions(arr)
    If dimensions = 1 Then
        Call swapItems_1D(arr, row1, row2)
    ElseIf dimensions = 2 Then
        Call swapItems_2D(arr, row1, row2)
    End If
End Sub

Private Sub swapItems_1D(arr As Variant, row1 As Long, row2 As Long)
    Dim tempItem As Variant
    '------------------------------------------------------------------------------------------------------
    Call f.Variables.assign(tempItem, arr(row2))
    Call f.Variables.assign(arr(row2), arr(row1))
    Call f.Variables.assign(arr(row1), tempItem)
End Sub

Private Sub swapItems_2D(arr As Variant, row1 As Long, row2 As Long)
    Dim tempArr() As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    ReDim tempArr(LBound(arr, 1) To UBound(arr, 1))
    For i = LBound(arr, 1) To UBound(arr, 1)
        Call f.Variables.assign(tempArr(i), arr(i, row2))
    Next i

    For i = LBound(arr, 1) To UBound(arr, 1)
        Call f.Variables.assign(arr(i, row2), arr(i, row1))
    Next i

    For i = LBound(arr, 1) To UBound(arr, 1)
        Call f.Variables.assign(arr(i, row1), tempArr(i))
    Next i

End Sub



Public Function invert(arr As Variant, Optional hasHeader As Boolean = False, Optional ByVal dimensions As Long = 0) As Variant
    If dimensions = 0 Then dimensions = f.arrays.countDimensions(arr)
    If dimensions = 1 Then
        invert = invert_1D(arr, hasHeader)
    ElseIf dimensions = 2 Then
        invert = invert_2D(arr, hasHeader)
    End If
End Function

Public Function invert_1D(arr As Variant, Optional hasHeader As Boolean = False) As Variant
    Dim i As Long
    Dim result() As Variant
    Dim boundsSum As Long
    Dim firstRow As Long
    '------------------------------------------------------------------------------------------------------

    firstRow = LBound(arr) + VBA.IIf(hasHeader, 1, 0)
    boundsSum = UBound(arr) + firstRow

    ReDim result(LBound(arr) To UBound(arr))
    If hasHeader Then result(LBound(result)) = arr(LBound(arr))

    For i = UBound(arr) To firstRow Step -1
        Call f.Variables.assign(result(i), arr(boundsSum - i))
    Next i
    invert_1D = result

End Function

Public Function invert_2D(arr As Variant, Optional hasHeader As Boolean = False) As Variant
    Dim i As Long
    Dim j As Long
    Dim boundsSum As Long
    Dim result() As Variant
    Dim firstRow As Long
    '------------------------------------------------------------------------------------------------------

    firstRow = LBound(arr, 2) + VBA.IIf(hasHeader, 1, 0)
    boundsSum = UBound(arr, 2) + firstRow

    ReDim result(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
    If hasHeader Then
        For j = LBound(arr, 1) To UBound(arr, 1)
            result(j, LBound(result)) = arr(j, LBound(arr))
        Next j
    End If

    For i = UBound(arr, 2) To firstRow Step -1
        For j = LBound(arr, 1) To UBound(arr, 1)
            Call f.Variables.assign(result(j, i), arr(j, boundsSum - i))
        Next j
    Next i
    invert_2D = result

End Function



Public Function getColumnIndexByHeader(arr As Variant, Header As Variant, Optional absoluteIndex As Boolean = False) As Long
    Dim headers As Variant
    '------------------------------------------------------------------------------------------------------
    headers = f.arrays.getColumn(arr, 1, True)
    getColumnIndexByHeader = getItemIndex(headers, Header)
End Function

Public Function getRowIndexByFirstValue(arr As Variant, value As Variant, Optional absoluteIndex As Boolean = False) As Long
    Dim values As Variant
    '------------------------------------------------------------------------------------------------------
    values = f.arrays.getRow(arr, 1, True)
    getRowIndexByFirstValue = getItemIndex(values, value)
End Function


Public Function getItemIndex(arr As Variant, item As Variant, Optional useRegex As Boolean = False, _
                                            Optional ByVal startColumn As Long = 123456789) As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If startColumn = 123456789 Then startColumn = LBound(arr)
    
    If Not f.arrays.isDefined(arr) Then
        ' skip - not an array
    ElseIf Not f.arrays.countDimensions(arr) = 1 Then
        ' skip - only 1D arrays are handled
    Else
        getItemIndex = LBound(arr) - 1
        For i = startColumn To UBound(arr)
            If useRegex Then
                If f.regex.checkIfMatch(VBA.CStr(arr(i)), VBA.CStr(item)) Then
                    getItemIndex = i
                    Exit For
                End If
            Else
                If f.Variables.equals(arr(i), item) Then
                    getItemIndex = i
                    Exit For
                End If
            End If
        Next i
    End If
    
End Function




Public Function transpose(ByRef arr As Variant) As Variant()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lWidth As Long
    Dim lHeight As Long
    Dim stWidth As Long
    Dim stHeight As Long
    Dim tempArray() As Variant
    Dim dimensions As Integer
    '------------------------------------------------------------------------------------------------------

    If Not f.arrays.isDefined(arr) Then Call Exceptions.NotArray.raise

    dimensions = f.arrays.countDimensions(arr)
    If dimensions <> 2 Then Call Exceptions.InvalidDimension.raise

    stWidth = LBound(arr, 1)
    stHeight = LBound(arr, 2)
    lWidth = UBound(arr, 1)
    lHeight = UBound(arr, 2)

    'Temporary array [tempArray] is given the target size.
    ReDim tempArray(stHeight To lHeight, stWidth To lWidth)
    For lngRow = stWidth To lWidth
        For lngCol = stHeight To lHeight

            If VBA.IsObject(arr(lngRow, lngCol)) Then
                Set tempArray(lngCol, lngRow) = arr(lngRow, lngCol)
            Else
                tempArray(lngCol, lngRow) = arr(lngRow, lngCol)
            End If

        Next lngCol
    Next lngRow

    'Newly created array [tempArray] is being F.Variables.assigned to the result variable.
    transpose = tempArray

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
NotArrayException:
    GoTo ExitPoint

DimensionsException:
    GoTo ExitPoint

End Function

Public Function transposeArray(ByRef arr As Variant) As Variant
    transposeArray = transpose(arr)
End Function



Public Sub expand(ByRef arr As Variant, Optional Rows As Long = 1)
    If f.arrays.countDimensions(arr) = 2 Then
        ReDim Preserve arr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2) + Rows)
    End If
End Sub

Public Function isNonEmptyArray(arr As Variant) As Boolean
    Dim dimensions As Integer
    '------------------------------------------------------------------------------------------------------
    If f.arrays.isDefined(arr) Then
        dimensions = f.arrays.countDimensions(arr)
        Select Case dimensions
            Case 1: isNonEmptyArray = isNonEmptyArray_1D(arr)
            Case 2: isNonEmptyArray = isNonEmptyArray_2D(arr)
        End Select
    End If
End Function

Private Function isNonEmptyArray_1D(arr As Variant) As Boolean
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    For i = LBound(arr) To UBound(arr)
        If Not VBA.IsEmpty(arr(i)) Then
            isNonEmptyArray_1D = True
            Exit For
        End If
    Next i
End Function

Private Function isNonEmptyArray_2D(arr As Variant) As Boolean
    Dim i As Long
    Dim j As Long
    '------------------------------------------------------------------------------------------------------
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If Not VBA.IsEmpty(arr(i, j)) Then
                isNonEmptyArray_2D = True
                Exit For
            End If
        Next j
    Next i
End Function



Public Function hasAnyUndefined(arr As Variant) As Boolean
    Dim dimensions As Integer
    '------------------------------------------------------------------------------------------------------
    dimensions = f.arrays.countDimensions(arr)
    Select Case dimensions
        Case 1: hasAnyUndefined = hasAnyUndefined_1D(arr)
        Case 2: hasAnyUndefined = hasAnyUndefined_2D(arr)
    End Select
End Function

Private Function hasAnyUndefined_1D(arr As Variant) As Boolean
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    For i = LBound(arr) To UBound(arr)
        If f.Variables.isUndefined(arr(i)) Then
            hasAnyUndefined_1D = True
            Exit For
        End If
    Next i
End Function

Private Function hasAnyUndefined_2D(arr As Variant) As Boolean
    Dim i As Long
    Dim j As Long
    '------------------------------------------------------------------------------------------------------
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If f.Variables.isUndefined(arr(i, j)) Then
                hasAnyUndefined_2D = True
                Exit For
            End If
        Next j
    Next i
End Function



Public Function join(ParamArray arrays() As Variant) As Variant
    Dim arr() As Variant
    Dim colArrays As VBA.Collection
    Dim items As Long
    Dim varItem As Variant
    Dim i As Long
    Dim upBound As Long:                            upBound = 1
    '------------------------------------------------------------------------------------------------------

    Set colArrays = New VBA.Collection
    For Each varItem In arrays
        If f.arrays.countDimensions(varItem) = 1 Then
            items = items + f.arrays.getSize(varItem)
            Call colArrays.Add(varItem)
        End If
    Next varItem

    ReDim arr(upBound To items)
    For Each varItem In colArrays
        For i = LBound(varItem) To UBound(varItem)
            arr(upBound + i - LBound(varItem)) = varItem(i)
        Next i
        upBound = upBound + UBound(varItem) - LBound(varItem) + 1
    Next varItem

    join = arr

End Function

Public Function join2D(ParamArray arrays() As Variant) As Variant
    Dim arr() As Variant
    Dim colArrays As VBA.Collection
    Dim varContainer As Variant
    Dim varArray As Variant
    Dim minColumn As Long
    Dim maxColumn As Long
    Dim items As Long
    Dim i As Long
    Dim j As Long
    Dim upBound As Long:                            upBound = 1
    '------------------------------------------------------------------------------------------------------

    Set colArrays = New VBA.Collection
    For Each varContainer In arrays
        For Each varArray In f.Collections.Create(varContainer)
            If f.arrays.countDimensions(varArray) = 2 Then
                maxColumn = Excel.Application.WorksheetFunction.Max(maxColumn, UBound(varArray, 1))
                If minColumn = 0 Then
                    minColumn = LBound(varArray, 1)
                Else
                    minColumn = Excel.Application.WorksheetFunction.Min(minColumn, LBound(varArray, 1))
                End If
                items = items + f.arrays.getSize(varArray, 2)
                Call colArrays.Add(varArray)
            End If
        Next varArray
    Next varContainer

    ReDim arr(minColumn To maxColumn, upBound To items)
    For j = minColumn To maxColumn
        upBound = 1
        For Each varArray In colArrays
            If f.Numerics.isInRange(j, LBound(varArray, 1), UBound(varArray, 1)) Then
                For i = LBound(varArray, 2) To UBound(varArray, 2)
                    arr(j, upBound + i - LBound(varArray)) = varArray(j, i)
                Next i
            End If
            upBound = upBound + UBound(varArray, 2) - LBound(varArray, 2) + 1
        Next varArray
    Next j

    join2D = arr

End Function


Public Function toDictionary(arr As Variant, Optional keyColumn As Long = 1, Optional valueColumn As Long = 2, _
                             Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Skip, _
                             Optional keyFunction As String = vbNullString, Optional keyFunctionInvoker As Object) _
                                                                                    As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Dim dimensions As Long
    Dim i As Long
    Dim key As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set toDictionary = f.dictionaries.Create(False)
    
    If f.arrays.isDefined(arr) Then
        dimensions = f.arrays.countDimensions(arr)
        If dimensions > 2 Then Call Exceptions.InvalidDimension.raiseWithCustomMessage("Function supports only 1D and 2D arrays")
        For i = LBound(arr, dimensions) To UBound(arr, dimensions)
            Select Case dimensions
                Case 1
                    Call f.Variables.assign(key, f.Reflection.modifyValue(arr(i), keyFunction, keyFunctionInvoker))
                    Call f.Variables.assign(value, arr(i))
                Case 2
                    Call f.Variables.assign(key, f.Reflection.modifyValue(arr(keyColumn, i), keyFunction, keyFunctionInvoker))
                    Call f.Variables.assign(value, arr(valueColumn, i))
            End Select
        
            With toDictionary
                If Not .exists(key) Then
                    Call .Add(key, value)
                ElseIf duplicateBehaviour = duplicateBehaviour_Override Then
                    If VBA.IsObject(value) Then
                        Set .item(key) = value
                    Else
                        .item(key) = value
                    End If
                ElseIf duplicateBehaviour = duplicateBehaviour_ThrowError Then
                    Call Exceptions.DuplicateKey.raise(f.strings.stringify(key))
                End If
            End With
        Next i
    End If
    
End Function

Public Function getArrayInfo(arr As Variant) As String
    Stop
End Function

Public Function getDescription(arr As Variant) As String
    Const EMPTY_ARRAY_PATTERN As String = "#EmptyArray [{0}]"
    Const INITIALIZED_ARRAY_PATTERN As String = "#Array [{0}] ({1})"
    '------------------------------------------------------------------------------------------------------
    Dim dimensions As Long
    Dim sizes() As String
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If f.arrays.isDefined(arr) Then
        dimensions = countDimensions(arr)
        ReDim sizes(1 To dimensions)
        
        For i = 1 To dimensions
            sizes(i) = LBound(arr, i) & " To " & UBound(arr, i)
        Next i
        getDescription = f.strings.format(INITIALIZED_ARRAY_PATTERN, VBA.TypeName(arr), VBA.join(sizes, ", "))
    Else
        getDescription = f.strings.format(EMPTY_ARRAY_PATTERN, VBA.TypeName(arr))
    End If
    
End Function

Public Function fromRange(rng As Excel.Range, Optional ifPossibleTo1D As Boolean = True) As Variant
    Dim arr As Variant
    Dim arrSheetIndices As Variant
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(rng) Then
        If rng.cells.count = 1 Then
            If ifPossibleTo1D Then
                fromRange = VBA.Array(rng.value)
            Else
                fromRange = to2D(VBA.Array(rng.value))
            End If
        ElseIf rng.Rows.count = 1 And ifPossibleTo1D Then
            arr = rng
            fromRange = f.arrays.getRow(arr, 1, False)
        ElseIf rng.columns.count = 1 And ifPossibleTo1D Then
            arr = rng
            fromRange = f.arrays.getColumn(arr, 1, False)
        Else
            fromRange = rng
        End If
    End If
    
End Function



Public Function fromParamArray(ParamArray items() As Variant) As Variant
    Dim arr As Variant
    Dim isFirstItemArray As Boolean
    Dim isParamArray As Boolean
    Dim varItem As Variant
    Dim Counter As Long
    '------------------------------------------------------------------------------------------------------
    
    'Only one item should be passed to this function (array of values). However to allow
    'passing ParamArray values directly to this function it has to be defined as ParamArray as well.
    For Each varItem In items
        arr = varItem
        Exit For
    Next varItem
    
    If Not VBA.IsEmpty(arr) Then
    
        For Each varItem In arr
            Counter = Counter + 1
            If VBA.IsArray(varItem) Then isFirstItemArray = True
            If Counter > 1 Then Exit For
        Next varItem
        
        isParamArray = (Counter = 1 And isFirstItemArray)
        
        If isParamArray Then
            For Each varItem In items
                fromParamArray = varItem(LBound(varItem))
                Exit For
            Next varItem
        Else
            fromParamArray = arr
        End If
        
    End If
    
End Function


Public Function findLastNonEmptyIndex(arr As Variant) As Long
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If countDimensions(arr) = 1 Then
        findLastNonEmptyIndex = LBound(arr) - 1
        For i = UBound(arr) To LBound(arr) Step -1
            If Not f.Variables.isUndefined(arr(i), True, True, True) Then
                findLastNonEmptyIndex = i
                Exit For
            End If
        Next i
    Else
        Call Exceptions.InvalidDimension.raiseWithCustomMessage("1D array expected")
    End If
    
End Function

Public Function to1D(value As Variant) As Variant

    If VBA.IsArray(value) Then
        Select Case countDimensions(value)
            Case 0, 1:          to1D = value
            Case 2:             to1D = convertTo1DArray_From2DArray(value)
            Case Else:          Call Exceptions.InvalidDimension.raiseWithCustomMessage("Function doesn't support arrays with more than 2 dimensions")
        End Select
    ElseIf VBA.IsObject(value) Then
        If TypeOf value Is Scripting.Dictionary Then
            to1D = f.cast.toDictionary(value).items
        ElseIf TypeOf value Is VBA.Collection Then
            to1D = f.Collections.toArray(f.cast.toCollection(value))
        Else
            to1D = VBA.Array(value)
        End If
    Else
        to1D = VBA.Array(value)
    End If

End Function

Private Function convertTo1DArray_From2DArray(arr As Variant) As Variant
    If f.arrays.getSize(arr, 1) = 1 Then
        convertTo1DArray_From2DArray = f.arrays.getRow(arr, 1, False)
    ElseIf f.arrays.getSize(arr, 2) = 1 Then
        convertTo1DArray_From2DArray = f.arrays.getColumn(arr, 1, False)
    Else
        Stop
        'TODO: Return array of arrays.
    End If
End Function

Public Function hasDuplicates(arr As Variant, Optional column As Long) As Boolean
    Dim dict As Scripting.Dictionary
    Dim dimensions As Long
    Dim i As Long
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set dict = f.dictionaries.Create(False)
    dimensions = f.arrays.countDimensions(arr)
    
    If dimensions < 1 Then
        '[arr] not declared
    ElseIf dimensions > 2 Then
        'array with more than 2 dimensions are not supported
    ElseIf Not f.Math.isInRange(column, LBound(arr, 1), True, UBound(arr, 1), True) Then
        'column index out of bound
    Else
        For i = LBound(arr, dimensions) To UBound(arr, dimensions)
            If dimensions = 1 Then
                Call f.Variables.assign(value, arr(i))
            Else
                Call f.Variables.assign(value, arr(column, i))
            End If
            
            If dict.exists(value) Then
                hasDuplicates = True
                Exit For
            Else
                Call dict.Add(value, vbEmpty)
            End If
            
        Next i
    End If
    
End Function



Public Function flatArray(arr As Variant) As Variant
    Dim result() As Variant
    Dim col As VBA.Collection
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    Dim subArr As Variant
    Dim size As Long
    Dim Counter As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not VBA.IsArray(arr) Then
        flatArray = VBA.Array(arr)
    ElseIf Not f.arrays.isDefined(arr) Then
        flatArray = VBA.Array()
    ElseIf f.arrays.countDimensions(arr) > 1 Then
        flatArray = VBA.Array(arr)
    Else
        Set col = New VBA.Collection
        For i = LBound(arr) To UBound(arr)
            Call col.Add(flatArray(arr(i)))
        Next i
        
        'Count total items in all subarrays
        For Each subArr In col
            size = size + f.arrays.getSize(subArr)
        Next subArr
        
        'Create final array
        ReDim result(1 To size)
        For Each subArr In col
            For i = LBound(subArr) To UBound(subArr)
                Counter = Counter + 1
                result(Counter) = subArr(i)
            Next i
        Next subArr
        
        flatArray = result
        
    End If
    
End Function


Public Function getColumnsMap(arr As Variant, Optional columnsInSecondDimension As Boolean = True) As Scripting.Dictionary
    Dim i As Long
    Dim Header As String
    Dim rowsDimension As Long:              rowsDimension = VBA.IIf(columnsInSecondDimension, 1, 2)
    Dim columnsDimension As Long:           columnsDimension = VBA.IIf(columnsInSecondDimension, 2, 1)
    '------------------------------------------------------------------------------------------------------
    
    If Not VBA.IsArray(arr) Then Call Exceptions.NotArray.raise
    If f.arrays.countDimensions(arr) <> 2 Then Call Exceptions.InvalidDimension.raise
    
    Set getColumnsMap = f.dictionaries.Create(False)
    
    For i = LBound(arr, columnsDimension) To UBound(arr, columnsDimension)
        If columnsInSecondDimension Then
            Header = arr(LBound(arr, columnsDimension), i)
        Else
            Header = arr(i, LBound(arr, columnsDimension))
        End If
        Call f.dictionaries.addItem(getColumnsMap, Header, i, duplicateBehaviour_Skip)
    Next i
    
End Function



Public Function getColumnIndex(arr As Variant, Header As Variant, _
                                                    Optional absoluteIndex As Boolean = False, _
                                                    Optional searchOnlyHeader As Boolean = True, _
                                                    Optional searchAsRegex As Boolean = False, _
                                                    Optional ByVal startColumn As Long = 123456789) As Long
    Dim headers As Variant
    Dim i As Long
    Dim j As Long
    Dim value As Variant
    Dim match As Boolean
    Dim objRegex As Object
    '------------------------------------------------------------------------------------------------------
    
    If Not f.arrays.isInBounds(startColumn, 1, arr) Then startColumn = LBound(arr, 1)
    
    If searchAsRegex Then Set objRegex = f.regex.Create(VBA.CStr(Header))
    
    If searchOnlyHeader Then
        headers = f.arrays.getColumn(arr, 1, False)
        getColumnIndex = getItemIndex(headers, Header, searchAsRegex, startColumn)
    Else
        For i = startColumn To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                value = arr(i, j)
                If VBA.IsError(value) Then
                    'Error from Excel cell - skip
                ElseIf searchAsRegex Then
                    match = objRegex.test(value)
                Else
                    match = (value = Header)
                End If
                
                If match Then
                    getColumnIndex = i
                    Exit For
                End If
            Next j
            If getColumnIndex Then Exit For
        Next i
    End If
                                                    
End Function


Public Function getRowIndex(arr As Variant, label As Variant, _
                            Optional absoluteIndex As Boolean = False, _
                            Optional searchOnlyFirstColumn As Boolean = True, _
                            Optional searchAsRegex As Boolean = False, _
                            Optional ByVal startRow As Long = 123456789) As Long
    Dim labels As Variant
    Dim i As Long
    Dim j As Long
    Dim value As Variant
    Dim match As Boolean
    Dim objRegex As Object
    '------------------------------------------------------------------------------------------------------
    
    If Not f.arrays.isInBounds(startRow, 1, arr) Then startRow = LBound(arr, 2)
    
    If searchAsRegex Then Set objRegex = f.regex.Create(VBA.CStr(label))
    
    If searchOnlyFirstColumn Then
        labels = f.arrays.getRow(arr, 1, True)
        getRowIndex = getItemIndex(labels, label, searchAsRegex)
    Else
        For i = LBound(arr, 2) To UBound(arr, 1)
            For j = startRow To UBound(arr, 2)
                value = arr(i, j)
                If searchAsRegex Then
                    match = objRegex.test(value)
                Else
                    match = (value = label)
                End If
                
                If match Then
                    getRowIndex = j
                    Exit For
                End If
            Next j
            If getRowIndex Then Exit For
        Next i
    End If
    
End Function














'***** WITHOUT UNIT TESTS *****


Public Function slice(arr As Variant, startIndex As Long, length As Long, Optional ByVal dimensions As Long = 0) As Variant
    If dimensions = 0 Then dimensions = f.arrays.countDimensions(arr)
    If dimensions = 1 Then
        slice = slice_1D(arr, startIndex, length)
    ElseIf dimensions = 2 Then
        slice = slice_2D(arr, startIndex, length)
    End If
End Function

Private Function slice_1D(arr As Variant, startIndex As Long, length As Long) As Variant
    Dim i As Long
    Dim j As Long
    Dim result() As Variant
    '------------------------------------------------------------------------------------------------------
    ReDim result(1 To length)
    For i = 1 To length
        Call f.Variables.assign(result(i), arr(i + startIndex - 1))
    Next i
    slice_1D = result
End Function

Private Function slice_2D(arr As Variant, startIndex As Long, length As Long) As Variant
    Dim i As Long
    Dim j As Long
    Dim result() As Variant
    '------------------------------------------------------------------------------------------------------
    ReDim result(LBound(arr, 1) To UBound(arr, 1), 1 To length)
    For i = 1 To length
        For j = LBound(arr, 1) To UBound(arr, 1)
            Call f.Variables.assign(result(j, i), arr(j, i + startIndex - 1))
        Next j
    Next i
    slice_2D = result
End Function



Public Function find2DArrayFieldByRegex(arr As Variant, regexPattern As String, Optional startFromEnd As Boolean = False) As Variant
    Dim objRegex As Object
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    Dim valueFound As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If countDimensions(arr) = 2 Then
        Set objRegex = f.regex.Create(regexPattern)
        If startFromEnd Then
            For i = UBound(arr, 1) To LBound(arr, 1) Step -1
                For j = UBound(arr, 2) To LBound(arr, 2) Step -1
                    If objRegex.test(VBA.CStr(arr(i, j))) Then
                        find2DArrayFieldByRegex = arr(i, j)
                        valueFound = True
                        Exit For
                    End If
                Next j
                If valueFound Then Exit For
            Next i
        Else
            For i = LBound(arr, 1) To UBound(arr, 1) Step -1
                For j = LBound(arr, 2) To UBound(arr, 2) Step -1
                    If objRegex.test(VBA.CStr(arr(i, j))) Then
                        find2DArrayFieldByRegex = arr(i, j)
                        valueFound = True
                        Exit For
                    End If
                Next j
                If valueFound Then Exit For
            Next i
        End If
        
    End If
End Function


Public Function toDtos(arr As Variant, Optional columns As Scripting.Dictionary) As VBA.Collection
    
    Set toDtos = New VBA.Collection
    
    If Not f.arrays.isDefined(arr) Then
        Call Exceptions.NotArray.raise
    ElseIf f.arrays.countDimensions(arr) <> 2 Then
        Call Exceptions.InvalidDimension.raiseWithCustomMessage("Function supports only 2D arrays")
    ElseIf f.dictionaries.isNonEmptyDictionary(columns) Then
        Set toDtos = getDtosWithPreselectedColumns(arr, columns)
    Else
        Set toDtos = getDtosArrayWithAllColumns(arr)
    End If
    
End Function

Private Function getDtosArrayWithAllColumns(arr As Variant) As VBA.Collection
    Dim headers As Variant
    Dim i As Long
    Dim j As Long
    '------------------------------------------------------------------------------------------------------
    Dim dto As Scripting.Dictionary
    Dim key As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------

    Set getDtosArrayWithAllColumns = New VBA.Collection

    headers = f.arrays.getRow(arr, 1, False)
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        Set dto = f.dictionaries.Create(False)
        For j = LBound(arr, 2) To UBound(arr, 2)
            key = headers(j)
            value = arr(i, j)
            Call f.dictionaries.addItem(dto, key, value, duplicateBehaviour_Override)
        Next j
        Call getDtosArrayWithAllColumns.Add(dto)
    Next i

End Function

Private Function getDtosWithPreselectedColumns(arr As Variant, columns As Scripting.Dictionary) As VBA.Collection
    Dim i As Long
    Dim colIndex As Long
    '------------------------------------------------------------------------------------------------------
    Dim dto As Scripting.Dictionary
    Dim key As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------

    Set getDtosWithPreselectedColumns = New VBA.Collection

    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        Set dto = f.dictionaries.Create(False)
        For Each key In columns.Keys
            colIndex = columns.item(key)
            value = arr(i, colIndex)
            Call f.dictionaries.addItem(dto, key, value, duplicateBehaviour_Override)
        Next key
        Call getDtosWithPreselectedColumns.Add(dto)
    Next i

End Function


Public Function toDtosDictionary(arr As Variant, keyProp As enumProperty, _
                                        Optional columns As Scripting.Dictionary) As Scripting.Dictionary
    Dim dtos As VBA.Collection
    Dim dto As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set toDtosDictionary = f.dictionaries.Create(False)
    Set dtos = toDtos(arr, columns)
    If f.Collections.isNonEmptyCollection(dtos) Then
        For Each dto In dtos
            key = f.dictionaries.getItem(dto, keyProp.getName)
            Call f.dictionaries.addItem(toDtosDictionary, key, dto)
        Next dto
    End If
    
End Function

Public Function getArraysOffset(arr1 As Variant, arr2 As Variant, Optional dimension As Long = 1) As Long
    getArraysOffset = LBound(arr1, dimension) - LBound(arr2, dimension)
End Function

Public Function getItemsSum(arr As Variant) As Double
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    If countDimensions(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            If VBA.isNumeric(arr(i)) Then
                getItemsSum = getItemsSum + arr(i)
            End If
        Next i
    End If
End Function









'[TRANSFERING DATA BETWEEN ARRAYS]

'[From old version - to be checked]
Public Sub copyContent(ByRef sourceArray As Variant, ByRef destinationArray As Variant)

    If Not VBA.IsArray(sourceArray) Then Call Exceptions.NotArray.raise

    If f.Variables.isDeclaredAsVariant(destinationArray) Then
        'If variable to which content is to be copied is declared as Variant,
        'it can be simply replace with the original array.
        destinationArray = sourceArray
    Else
        If Not VBA.IsArray(destinationArray) Then Call Exceptions.NotArray.raise
        If isDynamic(destinationArray) Then
            Call copyContentIntoDynamicArray(sourceArray, destinationArray)
        Else
            Call copyContentIntoStaticArray(sourceArray, destinationArray)
        End If
    End If
End Sub

Private Sub copyContentIntoDynamicArray(ByRef sourceArray As Variant, ByRef destinationArray As Variant)

    On Error GoTo ExceptionHandler
    destinationArray = sourceArray

'==========================================================================================================
ExitPoint:
    Exit Sub
'----------------------------------------------------------------------------------------------------------
ExceptionHandler:
    Call composeDynamicArrayFromOtherArray(sourceArray, destinationArray)
    GoTo ExitPoint
End Sub

Private Sub composeDynamicArrayFromOtherArray(ByRef sourceArray As Variant, ByRef destinationArray As Variant)
    Dim arrBounds() As Long
    '------------------------------------------------------------------------------------------------------
    arrBounds = getDimensionsBounds(sourceArray)
    Call resizeDynamicArrayToTheGivenBounds(destinationArray, arrBounds)
    Call populateArrayWithValuesFromOtherArray(sourceArray, destinationArray)
End Sub

Private Sub resizeDynamicArrayToTheGivenBounds(ByRef arr As Variant, bounds() As Long)
    Dim dimensions As Integer
    '------------------------------------------------------------------------------------------------------
    dimensions = getSize(bounds, 1)
    Select Case dimensions
        Case 0:         Erase arr
        Case 1:         ReDim arr(bounds(1, 1) To bounds(1, 2))
        Case 2:         ReDim arr(bounds(1, 1) To bounds(1, 2), _
                                  bounds(2, 1) To bounds(2, 2))
        Case Else:      Call Exceptions.InvalidDimension.raise
    End Select
End Sub

Private Sub populateArrayWithValuesFromOtherArray(ByRef sourceArray As Variant, ByRef destinationArray As Variant)
    Dim dimensions As Integer
    '------------------------------------------------------------------------------------------------------
    dimensions = countDimensions(sourceArray)
    Select Case dimensions
        Case 0:
        Case 1:         Call populateArrayWithValuesFrom1DArray(sourceArray, destinationArray)
        Case 2:         Call populateArrayWithValuesFrom2DArray(sourceArray, destinationArray)
        Case Else:      Call Exceptions.InvalidDimension.raise
    End Select
End Sub

Private Sub populateArrayWithValuesFrom1DArray(ByRef sourceArray As Variant, ByRef destinationArray As Variant)
    Dim item As Long
    Dim indexationDifference As Long
    '------------------------------------------------------------------------------------------------------
    indexationDifference = LBound(destinationArray) - LBound(sourceArray)
    For item = LBound(sourceArray) To UBound(sourceArray)
        Call f.Variables.assign(destinationArray(item + indexationDifference), sourceArray(item))
    Next item
End Sub

Private Sub populateArrayWithValuesFrom2DArray(ByRef sourceArray As Variant, ByRef destinationArray As Variant)
    Dim row As Long
    Dim column As Long
    Dim rowsIndexationDifference As Long
    Dim columnsIndexationDifference As Long
    '------------------------------------------------------------------------------------------------------

    rowsIndexationDifference = LBound(destinationArray, 1) - LBound(sourceArray, 1)
    columnsIndexationDifference = LBound(destinationArray, 2) - LBound(sourceArray, 2)

    For row = LBound(sourceArray) To UBound(sourceArray)
        For column = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            Call f.Variables.assign(destinationArray(row + rowsIndexationDifference, column + columnsIndexationDifference), _
                        sourceArray(row, column))
        Next column
    Next row
End Sub

Private Sub copyContentIntoStaticArray(ByRef sourceArray As Variant, ByRef destinationArray As Variant)
    If Not compareSizes(sourceArray, destinationArray, True) Then
        Call Exceptions.IndexOutOfBound.raise
    Else
        Call populateArrayWithValuesFromOtherArray(sourceArray, destinationArray)
    End If
End Sub

