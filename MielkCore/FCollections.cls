VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FCollections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FCollections"
'----------------------------------------------------------------------------------------------------------





'[CREATING COLLECTIONS]
Public Function Create(ParamArray items() As Variant) As VBA.Collection
    Dim item As Variant
    Dim hasContainer As Boolean
    '------------------------------------------------------------------------------------------------------
    
    If UBound(items) - LBound(items) = 0 Then
        If f.Variables.isContainer(items(LBound(items))) Then hasContainer = True
    End If
    
    If hasContainer Then
        Set Create = createFrom(items(LBound(items)))
    Else
        Set Create = New VBA.Collection
        For Each item In items
            Call Create.Add(item)
        Next item
    End If
    
End Function

Public Function createFrom(container As Variant) As VBA.Collection
    If f.Variables.isUndefined(container) Then
        Set createFrom = New VBA.Collection
    ElseIf f.arrays.countDimensions(container) = 1 Then
        Set createFrom = fromArray(container)
    ElseIf Not VBA.IsObject(container) Then
        Set createFrom = Create(container)
    ElseIf TypeOf container Is VBA.Collection Then
        Set createFrom = clone(f.cast.toCollection(container))
    ElseIf TypeOf container Is Scripting.Dictionary Then
        Set createFrom = f.dictionaries.toCollection(f.cast.toDictionary(container), DictPart_ValueOnly)
    Else
        Set createFrom = Create(container)
    End If
End Function

Public Function fromArray(arr As Variant) As VBA.Collection
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    Set fromArray = New VBA.Collection
    If f.arrays.countDimensions(arr) = 1 Then
        For i = LBound(arr) To UBound(arr)
            Call fromArray.Add(arr(i))
        Next i
    End If
End Function






'[CHECKING DATA TYPE]
Public Function isCollection(value As Variant) As Boolean
    If VBA.IsObject(value) Then
        If Not value Is Nothing Then
            isCollection = TypeOf value Is VBA.Collection
        End If
    End If
End Function

Public Function isNonEmptyCollection(value As Variant) As Boolean
    If isCollection(value) Then
        isNonEmptyCollection = (value.count > 0)
    End If
End Function

Public Function isEmptyCollection(value As Variant) As Boolean
    If isCollection(value) Then
        isEmptyCollection = (value.count = 0)
    End If
End Function











'[CONVERTING COLLECTION INTO OTHER DATA TYPES]

'#todo - add [deepClone] parameter
Public Function clone(col As VBA.Collection) As VBA.Collection
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    If col Is Nothing Then
    Else
        Set clone = New VBA.Collection
        For Each item In col
            Call clone.Add(item)
        Next item
    End If
    
End Function

'[fastModeForStringCollection] - use if you are sure collection contains only primitive values and you want it in plain format.
Public Function toString(col As VBA.Collection, _
                                    Optional ByVal stringifyMode As StringifyModeEnum = StringifyMode_Normal, _
                                    Optional ByVal itemsDelimiter As String = ", ", _
                                    Optional ByVal fastModeForStringCollection As Boolean = False, _
                                    Optional ByVal callbackMethod As String, _
                                    Optional ByVal callbackInvoker As Object, _
                                    Optional ByVal callbackParams As Scripting.Dictionary) As String
    Dim items() As String
    '------------------------------------------------------------------------------------------------------
    
    If col Is Nothing Then
        toString = f.strings.stringify(Nothing, stringifyMode)
    Else
        If fastModeForStringCollection Then
            items = convertStringCollectionToArray(col)
        Else
            items = toStringifedArray(col, stringifyMode, callbackMethod, callbackInvoker, callbackParams)
        End If
        toString = VBA.join(items, itemsDelimiter)
    End If
    
End Function

Private Function toStringifedArray(col As VBA.Collection, stringifyMode As StringifyModeEnum, _
                                    callbackMethod As String, callbackInvoker As Object, _
                                    callbackParams As Scripting.Dictionary) As String()
    Dim varItem As Variant
    Dim itemText As String
    Dim values() As String
    Dim counter As Long
    '------------------------------------------------------------------------------------------------------
    
    If col.count > 0 Then
        
        ReDim values(1 To col.count)
        
        For Each varItem In col
            counter = counter + 1
            If VBA.Len(callbackMethod) Then
                itemText = f.Reflection.modifyValue(varItem, callbackMethod, callbackInvoker, callbackParams)
            Else
                itemText = f.strings.stringify(varItem, stringifyMode)
            End If
            values(counter) = itemText
        Next varItem
        
        toStringifedArray = values
        
    End If

End Function

Private Function convertStringCollectionToArray(col As VBA.Collection) As String()
    Dim varItem As Variant
    Dim values() As String
    Dim counter As Long
    '------------------------------------------------------------------------------------------------------
    
    If col.count > 0 Then
        ReDim values(1 To col.count)
        For Each varItem In col
            counter = counter + 1
            values(counter) = varItem
        Next varItem
        convertStringCollectionToArray = values
    End If

End Function



'----------------------------------------------------------------------------------------------------------
' Description:                  Function to convert the given collection into dictionary.
' Returned value:               Dictionary based on the given collection. Each entry in the returned
'                               dictionary has the key and value equal to the corresponding item in collection.
'
' Parameters:
'   * col                       Source collection.
'   * duplicateBehaviour       (Optional) Defines what should happens if there are any duplicates in the
'                               original collection.
'   * keyFunction               (Optional) Function used to calculate key for each item in the original
'                               collection.
'                               If this parameter is missing, original value is used as both - key and value.
'                               If prefix $. is used in the name of function, collection items are treated
'                               as objects and their internal methods are used.
'   * keyFunctionInvoker        (Optional) Object used for invoking keyFunction.
'   * keyFunctionParams         _____
'----------------------------------------------------------------------------------------------------------
Public Function toDictionary(col As VBA.Collection, _
                    Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Skip, _
                    Optional keyFunction As String = vbNullString, Optional keyFunctionInvoker As Object, _
                    Optional keyFunctionParams As Scripting.Dictionary) As Scripting.Dictionary
    Const METHOD_NAME As String = "toDictionary"
    '------------------------------------------------------------------------------------------------------
    Dim key As Variant
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set toDictionary = f.dictionaries.Create(False)
    
    If isNonEmptyCollection(col) Then
        For Each item In col
            
            '[Calculate key]
            Call f.Variables.assign(key, f.Reflection.modifyValue(item, keyFunction, keyFunctionInvoker, keyFunctionParams))
            
            With toDictionary
                If Not .exists(key) Then
                    Call .Add(key, item)
                ElseIf duplicateBehaviour = duplicateBehaviour_ThrowError Then
                    Call Exceptions.DuplicateKey.raise(key)
                ElseIf duplicateBehaviour = duplicateBehaviour_Override Then
                    If VBA.IsObject(item) Then
                        Set .item(key) = item
                    Else
                        .item(key) = item
                    End If
                End If
            End With
        Next item
    End If
    
End Function


Public Function toArray(col As VBA.Collection, _
                        Optional callbackFunction As String = vbNullString, _
                        Optional callbackInvoker As Object) As Variant
    Dim arr() As Variant
    Dim index As Long
    '------------------------------------------------------------------------------------------------------
    
    If f.Collections.isNonEmptyCollection(col) Then
        ReDim arr(1 To col.count)
        For index = 1 To col.count
            Call f.Variables.assign(arr(index), _
                        f.Reflection.modifyValue(col.item(index), callbackFunction, callbackInvoker))
        Next index
        toArray = arr
    ElseIf f.Collections.isCollection(col) Then
        toArray = arr
    End If
    
End Function





'[COMPARING COLLECTIONS]
Public Function compare(baseCollection As VBA.Collection, _
                                    comparedCollection As VBA.Collection, _
                                    Optional strictDataType As Boolean = False, _
                                    Optional orderSensitive As Boolean = True, _
                                    Optional isCaseSensitive As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------

    If baseCollection Is Nothing Then
        compare = (comparedCollection Is Nothing)
    ElseIf comparedCollection Is Nothing Then
        compare = False
    ElseIf baseCollection.count <> comparedCollection.count Then
        compare = False
    ElseIf orderSensitive Then
        compare = compare_orderSensitive(baseCollection, comparedCollection, strictDataType, isCaseSensitive)
    Else
        compare = compare_orderInsensitive(baseCollection, comparedCollection, strictDataType, isCaseSensitive)
    End If
    
End Function

Private Function compare_orderSensitive( _
                            baseCollection As VBA.Collection, _
                            comparedCollection As VBA.Collection, _
                            strictDataType As Boolean, _
                            isCaseSensitive As Boolean) As Boolean
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    compare_orderSensitive = True
    For i = 1 To baseCollection.count
        If Not f.Variables.equals(baseCollection.item(i), comparedCollection.item(i), strictDataType, isCaseSensitive) Then
            compare_orderSensitive = False
            Exit For
        End If
    Next i
    
End Function

Private Function compare_orderInsensitive( _
                            baseCollection As VBA.Collection, _
                            comparedCollection As VBA.Collection, _
                            strictDataType As Boolean, _
                            isCaseSensitive As Boolean) As Boolean
    Dim baseCollectionClone As VBA.Collection
    Dim comparedCollectionClone As VBA.Collection
    Dim item As Variant
    Dim index As Long
    '------------------------------------------------------------------------------------------------------
    
    Set baseCollectionClone = f.Collections.clone(baseCollection)
    Set comparedCollectionClone = f.Collections.clone(comparedCollection)
    
    compare_orderInsensitive = True
    For Each item In baseCollectionClone
        index = f.Collections.getItemIndex(comparedCollectionClone, item, strictDataType)
        If index = 0 Then
            compare_orderInsensitive = False
            Exit For
        Else
            Call comparedCollectionClone.Remove(index)
        End If
    Next item
    
End Function

Public Function hasAnyCommonItemWithDictionary(col As VBA.Collection, dict As Scripting.Dictionary) As Boolean
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        For Each item In col
            If dict.exists(item) Then
                hasAnyCommonItemWithDictionary = True
                Exit For
            End If
        Next item
    End If
    
End Function





'[LOOKING FOR ITEMS]

Public Function getFirstItem(col As VBA.Collection) As Variant
    If isNonEmptyCollection(col) Then
        Call f.Variables.assign(getFirstItem, col.item(1))
    Else
        getFirstItem = Empty
    End If
End Function

Public Function getLastItem(col As VBA.Collection) As Variant
    If isNonEmptyCollection(col) Then
        Call f.Variables.assign(getLastItem, col.item(col.count))
    Else
        getLastItem = Empty
    End If
End Function

Public Function getLastObject(col As VBA.Collection) As Object
    If isNonEmptyCollection(col) Then
        Set getLastObject = col.item(col.count)
    Else
        Set getLastObject = Nothing
    End If
End Function

Public Function getItemByIndex(col As VBA.Collection, index As Long, Optional throwError As Boolean = False) As Variant
    If col Is Nothing Then
        If throwError Then Call Exceptions.NullNotAllowed.raise
    ElseIf index < 1 Or index > col.count Then
        If throwError Then Call Exceptions.IndexOutOfBound.raise
    Else
        getItemByIndex = col.item(index)
    End If
End Function

Public Function getObjectByIndex(col As VBA.Collection, index As Long, Optional throwError As Boolean = False) As Variant
    If col Is Nothing Then
        If throwError Then Call Exceptions.NullNotAllowed.raise
    ElseIf index < 1 Or index > col.count Then
        If throwError Then Call Exceptions.IndexOutOfBound.raise
    ElseIf Not VBA.IsObject(col.item(index)) Then
        If throwError Then Call Exceptions.ObjectExpected.raise
    Else
        Set getObjectByIndex = col.item(index)
    End If
End Function

Public Function contains(baseCollection As VBA.Collection, item As Variant, _
                                        Optional strictDataType As Boolean = False, _
                                        Optional isCaseSensitive As Boolean = False, _
                                        Optional trimIfValueIsString As Boolean = True) As Boolean
    Dim index As Long
    '------------------------------------------------------------------------------------------------------
    index = getItemIndex(baseCollection, item, strictDataType, isCaseSensitive, trimIfValueIsString)
    contains = (index > 0)
End Function

Public Function getItemIndex(baseCollection As VBA.Collection, item As Variant, _
                                        Optional strictDataType As Boolean = False, _
                                        Optional isCaseSensitive As Boolean = False, _
                                        Optional trimIfValueIsString As Boolean = True) As Long
    Dim collectionItem As Variant
    Dim index As Long
    '------------------------------------------------------------------------------------------------------

    If VBA.IsArray(item) Then Call Exceptions.IllegalDataType.raise

    For Each collectionItem In baseCollection
        index = index + 1
        If f.Variables.equals(item, collectionItem, strictDataType, isCaseSensitive, trimIfValueIsString) Then
            getItemIndex = index
            Exit For
        End If
    Next collectionItem

End Function

Public Function hasDuplicates(col As VBA.Collection, Optional modifyingFunction As String) As Boolean
    Dim dict As Scripting.Dictionary
    Dim item As Variant
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set dict = f.dictionaries.Create(False)
    
    If isNonEmptyCollection(col) Then
        For Each item In col
            Call f.Variables.assign(key, f.Reflection.modifyValue(item, modifyingFunction))
            If dict.exists(key) Then
                hasDuplicates = True
                Exit For
            Else
                Call dict.Add(key, vbNull)
            End If
        Next item
    End If
    
    Set dict = Nothing
    
End Function






'[APPENDING ITEMS]
Public Sub appendCollection(baseCollection As VBA.Collection, appendedCollection As VBA.Collection)
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    If appendedCollection Is Nothing Then
    ElseIf baseCollection Is Nothing Then
    Else
        For Each item In appendedCollection
            Call baseCollection.Add(item)
        Next item
    End If
End Sub

Public Sub appendItems(ByRef baseCollection As VBA.Collection, ParamArray items() As Variant)
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    If baseCollection Is Nothing Then Set baseCollection = New VBA.Collection
    For Each item In items
        Call appendItem(baseCollection, item, False)
    Next item
End Sub

Public Sub appendItemsWithFlatting(ByRef baseCollection As VBA.Collection, ParamArray items() As Variant)
    Dim item As Variant
    '----------------------------------------------------------------------------------------------------------
    If baseCollection Is Nothing Then Set baseCollection = New VBA.Collection
    For Each item In items
        Call appendItem(baseCollection, item, True)
    Next item
End Sub

Private Sub appendItem(baseCollection As VBA.Collection, item As Variant, Optional flatIfContainer As Boolean = True)
    Dim col As VBA.Collection
    Dim varItem As Variant
    '------------------------------------------------------------------------------------------------------
    
    If flatIfContainer And f.Variables.isContainer(item) Then
        Set col = f.Collections.Create(item)
        For Each varItem In col
            Call appendItem(baseCollection, varItem, flatIfContainer)
        Next varItem
    Else
        Call baseCollection.Add(item)
    End If
    
End Sub







Public Function join(ParamArray containers() As Variant) As VBA.Collection
    Dim container As Variant
    Dim key As Variant
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set join = New VBA.Collection
    
    For Each container In containers
        If VBA.IsArray(container) Then
            If f.arrays.countDimensions(container) = 1 Then
                For Each item In container
                    Call join.Add(item)
                Next item
            End If
        ElseIf VBA.IsObject(container) Then
            If isNonEmptyCollection(container) Then
                For Each item In container
                    Call join.Add(item)
                Next item
            ElseIf f.dictionaries.isNonEmptyDictionary(container) Then
                For Each key In container.Keys
                    Call join.Add(container.item(key))
                Next key
            End If
        End If
    Next container
    
End Function





'[REMOVING ITEMS]
Public Function removeDuplicates(col As VBA.Collection) As VBA.Collection
    Dim dict As Scripting.Dictionary
    Dim item As Variant
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set removeDuplicates = New VBA.Collection
    Set dict = f.dictionaries.Create(False)
    
    For Each item In col
        Call f.dictionaries.addItem(dict, item, item, duplicateBehaviour_Skip)
    Next item
    
    For Each key In dict.Keys
        Call removeDuplicates.Add(key)
    Next key
    
End Function

Public Sub removeItem(col As VBA.Collection, item As Variant)
    Dim i As Long
    Dim colItem As Variant
    '------------------------------------------------------------------------------------------------------
    
    If f.Collections.isNonEmptyCollection(col) Then
        For i = col.count To 1 Step -1
            Call f.Variables.assign(colItem, col.item(i))
            If VBA.IsObject(item) Then
                If VBA.IsObject(colItem) Then
                    If item Is colItem Then
                        Call col.Remove(i)
                    End If
                End If
            ElseIf Not VBA.IsArray(item) Then
                If Not VBA.IsObject(colItem) And Not VBA.IsArray(colItem) Then
                    If item = colItem Then
                        Call col.Remove(i)
                    End If
                End If
            End If
        Next i
    End If
    
End Sub





'[FILTERING & SORTING]
Public Function getUniqueItems(col As VBA.Collection) As VBA.Collection
    Dim dict As Scripting.Dictionary
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set dict = f.dictionaries.Create(False)
    With dict
        For Each item In col
            If Not .exists(item) Then
                Call .Add(item, item)
            End If
        Next item
    End With
    
    Set getUniqueItems = f.dictionaries.toCollection(dict)
    
End Function



'----------------------------------------------------------------------------------------------------------
' Description:          Method to filter the given collection using the given evaluation rule.
'----------------------------------------------------------------------------------------------------------
Public Function filterObjects(col As VBA.Collection, evaluationFunction As String) As VBA.Collection
    Dim item As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set filterObjects = New VBA.Collection
    For Each item In col
        If VBA.IsObject(item) Then
            If Not item Is Nothing Then
                If checkIfIncludeItem(f.cast.toObject(item), evaluationFunction) Then
                    Call filterObjects.Add(item)
                End If
            End If
        End If
    Next item
    
End Function


Private Function checkIfIncludeItem(item As Object, ByVal evaluationFunction As String) As Boolean
    Dim substr As String
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Do
        substr = f.strings.substring(evaluationFunction, METHOD_START_TAG, METHOD_END_TAG)
        If VBA.Len(substr) Then
            value = VBA.CallByName(item, substr, VbMethod)
            evaluationFunction = VBA.replace(evaluationFunction, METHOD_START_TAG & substr & METHOD_END_TAG, value)
        End If
    Loop While VBA.Len(substr)
    
    checkIfIncludeItem = Application.Evaluate(evaluationFunction)
    
End Function



Public Function sort(col As VBA.Collection, Optional ByVal ascending As Boolean = True, _
                        Optional sortFunction As String = vbNullString, _
                        Optional sortFunctionInvoker As Object) As VBA.Collection
    Dim arr As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    Set sort = New VBA.Collection
    
    If isNonEmptyCollection(col) Then
        ReDim arr(1 To 2, 1 To col.count)
        For i = LBound(arr, 2) To UBound(arr, 2)
            arr(1, i) = f.Reflection.modifyValue(col.item(i), sortFunction, sortFunctionInvoker)
            Call f.Variables.assign(arr(2, i), col.item(i))
        Next i
        arr = f.arrays.sort(arr, ascending, 1)
        For i = LBound(arr, 2) To UBound(arr, 2)
            Call sort.Add(arr(2, i))
        Next i
    End If

End Function


Public Function filter(col As VBA.Collection, _
                        Optional filterMethod As String = vbNullString, _
                        Optional filterMethodInvoker As Object, _
                        Optional filterMethodParams As Scripting.Dictionary) As VBA.Collection
    Dim item As Variant
    Dim include As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set filter = New VBA.Collection
    If Not col Is Nothing Then
        For Each item In col
            If VBA.Len(filterMethod) Then
                include = f.Reflection.modifyValue(item, filterMethod, filterMethodInvoker, filterMethodParams)
            Else
                include = False
            End If
            
            If include Then
                Call filter.Add(item)
            End If
            
        Next item
    End If
    

End Function

Public Function map(col As VBA.Collection, modifyMethod As String, _
                        Optional modifyMethodInvoker As Object, _
                        Optional modifyMethodParams As Scripting.Dictionary) As VBA.Collection
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set map = New VBA.Collection
    
    If isNonEmptyCollection(col) Then
        For Each item In col
            Call map.Add(f.Reflection.modifyValue(item, modifyMethod, modifyMethodInvoker, modifyMethodParams))
        Next item
    End If
    
End Function


Public Function forEach(col As VBA.Collection, methodName As String, _
                                Optional methodInvoker As Object, _
                                Optional methodParams As Scripting.Dictionary)
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    If isNonEmptyCollection(col) Then
        For Each item In col
            Call f.Reflection.runMethod(item, methodName, methodInvoker, methodParams)
        Next item
    End If
End Function


Public Function removeUndefined(col As VBA.Collection, Optional includeEmptyString As Boolean = False, _
                                    Optional includeZero As Boolean = False, _
                                    Optional includeUndefinedArray As Boolean = True) As VBA.Collection
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
                                    
    If isCollection(col) Then
        Set removeUndefined = New VBA.Collection
        For Each item In col
            If f.Variables.isUndefined(item, includeEmptyString, includeZero, includeUndefinedArray) Then
            Else
                Call removeUndefined.Add(item)
            End If
        Next item
    End If
                                    
End Function


Public Function slice(col As VBA.Collection, startIndex As Long, size As Long) As VBA.Collection
    Dim i As Long
    Dim endIndex As Long
    '------------------------------------------------------------------------------------------------------
    
    Set slice = New VBA.Collection
    If Not f.Collections.isNonEmptyCollection(col) Then
        '
    ElseIf startIndex > col.count Then
        '
    Else
        endIndex = f.Numerics.getMin(startIndex + size - 1, col.count)
        For i = startIndex To endIndex
            Call slice.Add(col.item(i))
        Next i
    End If
    
End Function


'
'
'
'
''***** WITHOUT UNIT TESTS *****
'
'
'
'Public Function appendItems(baseCollection As VBA.Collection, ParamArray items() As Variant) As VBA.Collection
'    Dim item As Variant
'    '------------------------------------------------------------------------------------------------------
'    If baseCollection Is Nothing Then Set baseCollection = New VBA.Collection
'    For Each item In items
'        Call appendItem(baseCollection, item)
'    Next item
'    Set appendItems = baseCollection
'End Function
'
'Private Sub appendItem(baseCollection As VBA.Collection, item As Variant, Optional flatContainers As Boolean = True)
'    Dim col As VBA.Collection
'    Dim varItem As Variant
'    '------------------------------------------------------------------------------------------------------
'
'    If flatContainers Then
'        If F.Variables.isContainer(item) Then
'            Set col = F.Collections.Create(item)
'            For Each varItem In col
'                Call appendItem(baseCollection, varItem, flatContainers)
'            Next varItem
'        Else
'            Call baseCollection.Add(item)
'        End If
'    Else
'        Call baseCollection.Add(item)
'    End If
'
'End Sub
'
'Public Function appendItemsWithoutFlatting(baseCollection As VBA.Collection, ParamArray items() As Variant) As VBA.Collection
'    Dim item As Variant
'    '------------------------------------------------------------------------------------------------------
'    If baseCollection Is Nothing Then Set baseCollection = New VBA.Collection
'    For Each item In items
'        Call appendItem(baseCollection, item, False)
'    Next item
'    Set appendItemsWithoutFlatting = baseCollection
'End Function
'
'Public Function removeUndefinedItems(items As VBA.Collection) As VBA.Collection
'    Dim item As Variant
'    '------------------------------------------------------------------------------------------------------
'    Set removeUndefinedItems = New VBA.Collection
'    For Each item In items
'        If Not F.Variables.isUndefined(item, True, True, True) Then
'            Call removeUndefinedItems.Add(item)
'        End If
'    Next item
'End Function
'


Public Function getMaxValue(col As VBA.Collection) As Variant
    Dim initialized As Boolean
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    If Not col Is Nothing Then
        For Each item In col
            If Not initialized Then
                getMaxValue = item
                initialized = True
            ElseIf item > getMaxValue Then
                getMaxValue = item
            End If
        Next item
    End If
End Function


Public Function fromString(value As String, Optional separator As String = ",") As VBA.Collection
    Dim arr As Variant
    Dim i As Long
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set fromString = New VBA.Collection
    
    arr = VBA.Split(value, separator)
    If f.arrays.isDefined(arr) Then
        For i = LBound(arr) To UBound(arr)
            item = arr(i)
            If VBA.isNumeric(item) Then item = item + 1
            Call fromString.Add(item)
        Next i
    End If
    
End Function


Public Sub append(baseCollection As VBA.Collection, appendedCollection As VBA.Collection)
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    If Not appendedCollection Is Nothing Then
        For Each item In appendedCollection
            Call baseCollection.Add(item)
        Next item
    End If
End Sub
