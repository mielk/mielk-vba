VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FDictionaries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FDictionaries"
'----------------------------------------------------------------------------------------------------------
Private Const DICTIONARY_TYPENAME As String = "Dictionary"
'----------------------------------------------------------------------------------------------------------




'[CREATING]
Public Function Create(Optional isCaseSensisitve As Boolean = False) As Object
    Set Create = New Scripting.Dictionary
    Create.CompareMode = VBA.IIf(isCaseSensisitve, DictCompareMode_Binary, DictCompareMode_Text)
End Function

Public Function createWithItems(isCaseSensitive As Boolean, ParamArray keyValuePairs() As Variant) As Scripting.Dictionary
    Dim varItem As Variant
    Dim key As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set createWithItems = Create(isCaseSensitive)
    
    For Each varItem In keyValuePairs
        With createWithItems
            If Not .exists(varItem(1)) Then
                Call .Add(varItem(1), varItem(2))
            Else
                Call Exceptions.DuplicateKey.raise(F.strings.stringify(varItem(1)))
            End If
        End With
    Next varItem
    
End Function

Public Function createFrom(container As Variant, _
                            Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Skip, _
                            Optional keyFunction As String = vbNullString, Optional keyFunctionInvoker As Object) _
                            As Scripting.Dictionary

    If F.Variables.isString(container) Then
        Set createFrom = fromString(VBA.CStr(container))
    ElseIf VBA.IsArray(container) Then
        Set createFrom = F.arrays.toDictionary(container, duplicateBehaviour:=duplicateBehaviour, _
                                                keyFunction:=keyFunction, keyFunctionInvoker:=keyFunctionInvoker)
    ElseIf Not F.Variables.isContainer(container) Then
        ' skip
    ElseIf F.Dictionaries.isDictionary(container) Then
        Set createFrom = container
    ElseIf F.Collections.isCollection(container) Then
        Set createFrom = F.Collections.toDictionary(F.cast.toCollection(container), duplicateBehaviour, keyFunction, keyFunctionInvoker)
    End If
    

End Function

Public Function fromString(ByVal text As String, _
                                    Optional ByVal entriesSeparator As String = ";", _
                                    Optional ByVal keyItemSeparator As String = ":", _
                                    Optional ByVal behaviourOnDuplicate As DuplicateBehaviourEnum = duplicateBehaviour_Override) As Scripting.Dictionary
    Const REGEX_PATTERN As String = "([^{0}{1}]+)(?:{0}?)([^{1}]*)"
    '------------------------------------------------------------------------------------------------------
    Dim regexPattern As String
    Dim entries As VBA.Collection
    Dim arr As Variant
    Dim key As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set fromString = F.Dictionaries.Create(False)
    
    regexPattern = F.strings.format(REGEX_PATTERN, keyItemSeparator, entriesSeparator)
    Set entries = F.regex.getMatchArraysCollection(text, regexPattern)
    For Each arr In entries
        key = arr(LBound(arr))
        value = arr(UBound(arr))
        If VBA.isNumeric(value) Then value = value * 1
        Call F.Dictionaries.addItem(fromString, key, value, behaviourOnDuplicate)
    Next arr
    
End Function





'[CHECKING DATA TYPE]
Public Function isDictionary(ByRef value As Variant) As Boolean
    isDictionary = (VBA.StrComp(VBA.TypeName(value), DICTIONARY_TYPENAME, vbTextCompare) = 0)
End Function

Public Function isEmptyDictionary(ByRef value As Variant) As Boolean
    If isDictionary(value) Then
        isEmptyDictionary = (value.count = 0)
    End If
End Function

Public Function isNonEmptyDictionary(ByRef value As Variant) As Boolean
    If isDictionary(value) Then
        isNonEmptyDictionary = (value.count > 0)
    End If
End Function






'[ACCESS TO ITEMS]
Public Function getItem(dict As Scripting.Dictionary, key As Variant) As Variant
    If Not dict Is Nothing Then
        With dict
            If .exists(key) Then
                If VBA.IsObject(.item(key)) Then
                    Set getItem = .item(key)
                Else
                    getItem = .item(key)
                End If
            End If
        End With
    End If
End Function

Public Function getObject(dict As Object, key As Variant) As Object
    If Not dict Is Nothing Then
        With dict
            If .exists(key) Then
                If VBA.IsObject(.item(key)) Then
                    Set getObject = .item(key)
                End If
            End If
        End With
    End If
End Function

Public Function getFirstEntry(dict As Scripting.Dictionary, _
                                        Optional partType As DictPartEnum = DictPart_ValueOnly) As Variant
    Dim key As Variant
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------
    
    For Each key In dict.Keys
        If partType = DictPart_KeyOnly Then
            Call F.Variables.assign(getFirstEntry, key)
        ElseIf partType = DictPart_ValueOnly Then
            Call F.Variables.assign(getFirstEntry, dict.item(key))
        Else
            getFirstEntry = KeyValue(key, dict.item(key))
        End If
        Exit For
    Next key
    
End Function


Public Sub assignValueFromDictToVariable(dict As Scripting.Dictionary, variable As Variant, tags As String, _
                                            Optional conversionFunction As String)
    Const TAGS_SEPARATOR As String = ";"
    '------------------------------------------------------------------------------------------------------
    Dim arrTags() As String
    Dim value As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    arrTags = VBA.Split(tags, TAGS_SEPARATOR)
    If F.arrays.isDefined(arrTags) Then
        For i = LBound(arrTags) To UBound(arrTags)
            Call F.Variables.assign(value, getPropValue(dict, VBA.trim$(arrTags(i)), conversionFunction))
            If Not VBA.IsEmpty(value) Then
                Call F.Variables.assign(variable, value)
                Exit For
            End If
        Next i
    End If
    

End Sub

Public Function getPropValue(dict As Scripting.Dictionary, tag As String, Optional conversionFunction As String) As Variant
    Dim arrTags As Variant
    Dim i As Long
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    arrTags = VBA.Split(tag, ".")
    
    If F.Dictionaries.isNonEmptyDictionary(dict) Then
        If F.arrays.isDefined(arrTags) Then
            For i = LBound(arrTags) To UBound(arrTags)
                With dict
                    If .exists(arrTags(i)) Then
                        If i = UBound(arrTags) Then
                            If VBA.Len(conversionFunction) Then
                                On Error GoTo ConversionException
                                Call F.Variables.assign(getPropValue, Application.Run(conversionFunction, .item(arrTags(i))))
                                
                            Else
                                Call F.Variables.assign(getPropValue, .item(arrTags(i)))
                            End If
                        Else
                            Exit For
                        End If
                    End If
                End With
            Next i
        End If
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Function

'----------------------------------------------------------------------------------------------------------
ConversionException:
    Call VBA.err.raise(1, description:="Error when converting with function " & conversionFunction)

End Function






'[CLONING]
Public Function clone(dictSource As Scripting.Dictionary, Optional deepClone As Boolean = False) As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------

    If dictSource Is Nothing Then
        Set clone = Nothing
    Else
        Set clone = New Scripting.Dictionary
        clone.CompareMode = dictSource.CompareMode
        For Each key In dictSource.Keys
            Call clone_assignValue(clone, key, dictSource.item(key), deepClone)
        Next key
    End If

End Function

Private Sub clone_assignValue(dict As Scripting.Dictionary, key As Variant, value As Variant, Optional deepClone As Boolean = False)
    If Not deepClone Then
        Call dict.Add(key, value)
    ElseIf F.Dictionaries.isDictionary(value) Then
        Call clone_deepCloning(value, dict, key)
    ElseIf F.Collections.isCollection(value) Then
        Call clone_deepCloning(value, dict, key)
    Else
        Call dict.Add(key, value)
    End If
End Sub

Private Function clone_deepCloning(sourceContainer As Variant, destination As Variant, key As Variant) As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Dim col As VBA.Collection
    Dim varKey As Variant
    Dim varItem As Variant
    '------------------------------------------------------------------------------------------------------
    If F.Dictionaries.isNonEmptyDictionary(sourceContainer) Then
        Set dict = F.Dictionaries.Create(False)
        Call destination.Add(key, dict)
        For Each varKey In sourceContainer.Keys
            Call clone_assignValue(dict, varKey, sourceContainer.item(varKey), True)
        Next varKey
    ElseIf F.Collections.isNonEmptyCollection(sourceContainer) Then
        Set col = New VBA.Collection
        Call destination.Add(key, col)
        For Each varItem In sourceContainer
            Call clone_assignValue(col, vbNullString, varItem, True)
        Next varItem
    End If
End Function







'[SORTING]
Public Function sortByKey(dict As Scripting.Dictionary, Optional ascendingOrder As Boolean = True) As Scripting.Dictionary
    Dim arr As Variant
    Dim key As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If Not dict Is Nothing Then
        Set sortByKey = F.Dictionaries.Create(dict.CompareMode = TextCompare)
        If dict.count Then
            ReDim arr(1 To 2, 1 To dict.count)
            For Each key In dict.Keys
                i = i + 1
                Call F.Variables.assign(arr(1, i), key)
                Call F.Variables.assign(arr(2, i), dict.item(key))
            Next key
            
            Call F.arrays.sort(arr, ascendingOrder, 1, False)
            
            For i = LBound(arr, 2) To UBound(arr, 2)
                Call sortByKey.Add(arr(1, i), arr(2, i))
            Next i
            
        End If
    End If
    
End Function





'[CONVERTING COLLECTION INTO OTHER DATA TYPES]

Public Function toString(dict As Object, _
                                        Optional part As DictPartEnum = DictPart_KeyAndValue, _
                                        Optional stringifyMode As StringifyModeEnum = StringifyMode_Normal, _
                                        Optional itemsDelimiter As String = ", ", _
                                        Optional keyItemDelimiter As String = ":", _
                                        Optional entryWrapper As String = "{}") As String
    Dim entries() As String
    '------------------------------------------------------------------------------------------------------

    If dict Is Nothing Then
        toString = F.strings.stringify(Nothing, stringifyMode)
    ElseIf Not isDictionary(dict) Then
        Call Exceptions.IllegalDataType.raise
    Else
        entries = toString_getEntriesArray(dict, part, stringifyMode, keyItemDelimiter, entryWrapper)
        If stringifyMode = StringifyMode_Xml Then
            toString = F.strings.format(XML_DICTIONARY, VBA.join(entries, vbNullString))
        Else
            toString = VBA.join(entries, itemsDelimiter)
        End If
    End If

End Function

Private Function toString_getEntriesArray(dict As Object, part As DictPartEnum, _
                                        stringifyMode As StringifyModeEnum, _
                                        ByVal keyItemDelimiter As String, _
                                        ByVal entryWrapper As String) As String()
    Dim varKey As Variant
    Dim values() As String
    Dim counter As Long
    '------------------------------------------------------------------------------------------------------
    Dim openWrapper As String
    Dim closeWrapper As String
    '------------------------------------------------------------------------------------------------------

    If dict.count > 0 Then
    
        If stringifyMode <> StringifyMode_Xml And part = DictPart_KeyAndValue Then
            openWrapper = VBA.left$(entryWrapper, 1)
            closeWrapper = VBA.right$(entryWrapper, 1)
        End If

        ReDim values(1 To dict.count)
        For Each varKey In dict.Keys
            counter = counter + 1
            values(counter) = openWrapper & _
                                getItemText(dict, varKey, part, stringifyMode, keyItemDelimiter) & _
                                closeWrapper
        Next varKey

    End If

    toString_getEntriesArray = values

End Function

Private Function getItemText(dict As Object, varKey As Variant, part As DictPartEnum, _
                                        stringifyMode As StringifyModeEnum, _
                                        ByVal keyItemDelimiter As String) As String
    Const KEY_OPEN_TAG As String = "<key>"
    Const KEY_CLOSE_TAG As String = "</key>"
    Const ITEM_OPEN_TAG As String = "<item>"
    Const ITEM_CLOSE_TAG As String = "</item>"
    '------------------------------------------------------------------------------------------------------
    Dim openTag As String
    Dim closeTag As String
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------

    If stringifyMode = StringifyMode_Xml Then
        If part = DictPart_KeyOnly Then
            openTag = KEY_OPEN_TAG
            closeTag = KEY_CLOSE_TAG
        ElseIf part = DictPart_ValueOnly Then
            openTag = ITEM_OPEN_TAG
            closeTag = ITEM_CLOSE_TAG
        End If
    End If

    If part = DictPart_KeyOnly Then
        getItemText = openTag & F.strings.stringify(varKey, stringifyMode) & closeTag
    ElseIf part = DictPart_ValueOnly Then
        Call F.Variables.assign(value, dict.item(varKey))
        getItemText = openTag & F.strings.stringify(value, stringifyMode) & closeTag
    ElseIf part = DictPart_KeyAndValue Then
        Call F.Variables.assign(value, dict.item(varKey))
        getItemText = getItemText_KeyValue(varKey, value, stringifyMode, keyItemDelimiter)
    End If

End Function

Private Function getItemText_KeyValue(key As Variant, value As Variant, _
                                                stringifyMode As StringifyModeEnum, _
                                                ByVal keyItemDelimiter As String) As String
    Dim keyText As String
    '------------------------------------------------------------------------------------------------------

    If stringifyMode = StringifyMode_Xml Then
        keyText = F.strings.stringify(key, StringifyMode_Normal)
        getItemText_KeyValue = "<" & keyText & ">" & _
                                F.strings.stringify(value, stringifyMode) & _
                                "</" & keyText & ">"
    Else
        getItemText_KeyValue = F.strings.stringify(key, StringifyMode_Normal) & _
                                keyItemDelimiter & _
                                F.strings.stringify(value, stringifyMode)
    End If

End Function



Public Function toArray(dict As Scripting.Dictionary, Optional dictPart As DictPartEnum = DictPart_KeyAndValue) As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------

    If dict Is Nothing Then
        ' return Empty
    ElseIf dict.count Then
        toArray = arr
    Else
        If dictPart = DictPart_KeyAndValue Then
            toArray = F.arrays.to2D(dict)
        ElseIf dictPart = DictPart_KeyOnly Then
            toArray = dict.Keys
        ElseIf dictPart = DictPart_ValueOnly Then
            toArray = dict.items
        End If
        toArray = arr
    End If
    
End Function

Public Function toCollection(dict As Scripting.Dictionary, _
                                Optional partType As DictPartEnum = DictPart_ValueOnly) As VBA.Collection
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------

    Set toCollection = New VBA.Collection
    If isNonEmptyDictionary(dict) Then
        For Each key In dict.Keys
            If partType = DictPart_ValueOnly Then
                Call toCollection.Add(dict.item(key))
            ElseIf partType = DictPart_KeyOnly Then
                Call toCollection.Add(key)
            End If
        Next key
    End If

End Function











'[ADDING & REMOVING ITEMS]
Public Sub addItem(dict As Scripting.Dictionary, key As Variant, item As Variant, _
                    Optional behaviourIfDuplicate As DuplicateBehaviourEnum = duplicateBehaviour_Override)
                    
    If dict Is Nothing Then
        Call Exceptions.NullNotAllowed.raiseWithCustomMessage("Dictionary cannot be Nothing")
    End If
    
    With dict
        If .exists(key) Then
            Select Case behaviourIfDuplicate
                Case duplicateBehaviour_Skip
                Case duplicateBehaviour_Override
                    If VBA.IsObject(item) Then
                        Set .item(key) = item
                    Else
                        .item(key) = item
                    End If
                Case duplicateBehaviour_ThrowError
                    Call Exceptions.DuplicateKey.raise(key)
                Case duplicateBehaviour_WarningInImmediateWindow
                    Debug.Print "Duplicate key " & key
            End Select
        Else
            Call .Add(key, item)
        End If
    End With
    
End Sub

Public Sub appendItemsFromContainer(dict As Scripting.Dictionary, container As Variant, _
                    Optional behaviourIfDuplicate As DuplicateBehaviourEnum = duplicateBehaviour_Override)
    
    If Not F.Variables.isContainer(container) Then
        Call Exceptions.IllegalDataType.raiseWithCustomMessage("Given parameters [container] is not container")
    End If
    
    If Not dict Is Nothing Then
        If F.Dictionaries.isDictionary(container) Then
            Call appendDictionary(dict, F.cast.toDictionary(container), behaviourIfDuplicate)
        ElseIf F.Collections.isNonEmptyCollection(container) Then
            Call appendCollection(dict, F.cast.toCollection(container), behaviourIfDuplicate)
        ElseIf F.arrays.isDefined(container) Then
            Stop
        End If
    End If
    
End Sub

Public Sub appendDictionary(baseDict As Scripting.Dictionary, mergedDict As Scripting.Dictionary, _
                    Optional behaviourIfDuplicate As DuplicateBehaviourEnum = duplicateBehaviour_Override)
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If baseDict Is Nothing Then
        Call Exceptions.NullNotAllowed.raiseWithCustomMessage("Dictionary cannot be Nothing")
    End If
    
    If isNonEmptyDictionary(mergedDict) Then
        For Each key In mergedDict.Keys
            Call addItem(baseDict, key, mergedDict.item(key), behaviourIfDuplicate)
        Next key
    End If
    
End Sub

Public Sub appendCollection(baseDict As Scripting.Dictionary, col As VBA.Collection, _
                    Optional behaviourIfDuplicate As DuplicateBehaviourEnum = duplicateBehaviour_Override)
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    
    If baseDict Is Nothing Then
        Call Exceptions.NullNotAllowed.raiseWithCustomMessage("Dictionary cannot be Nothing")
    End If
    
    If F.Collections.isNonEmptyCollection(col) Then
        For Each item In col
            Call addItem(baseDict, item, item, behaviourIfDuplicate)
        Next item
    End If
    
End Sub

Public Sub appendArray(baseDict As Scripting.Dictionary, arr As Variant, _
                    Optional behaviourIfDuplicate As DuplicateBehaviourEnum = duplicateBehaviour_Override)
    Dim dimensions As Long
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If baseDict Is Nothing Then
        Call Exceptions.NullNotAllowed.raiseWithCustomMessage("Dictionary cannot be Nothing")
    End If
    
    dimensions = F.arrays.countDimensions(arr)
    If dimensions = 1 Then
        For i = LBound(arr) To UBound(arr)
            Call addItem(baseDict, arr(i), arr(i), behaviourIfDuplicate)
        Next i
    ElseIf dimensions = 2 Then
        '#TODO
        Stop
    End If
    
End Sub

Public Function merge(masterDict As Scripting.Dictionary, appendedDict As Scripting.Dictionary, _
                        Optional deepMerge As Boolean = True, _
                        Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Override) As Scripting.Dictionary
    Dim key As Variant
    Dim subDict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If Not masterDict Is Nothing Then
        Set merge = F.Dictionaries.clone(masterDict)
    Else
        Set merge = F.Dictionaries.Create(False)
    End If
    
    If Not appendedDict Is Nothing Then
        For Each key In appendedDict
            If Not merge.exists(key) Then
                Call merge.Add(key, appendedDict.item(key))
            ElseIf F.Dictionaries.isDictionary(merge.item(key)) And F.Dictionaries.isDictionary(appendedDict.item(key)) Then
                Set subDict = merge(merge.item(key), appendedDict.item(key), deepMerge, duplicateBehaviour)
                Set merge.item(key) = subDict
            Else
                Call F.Dictionaries.addItem(merge, key, appendedDict.item(key), duplicateBehaviour)
            End If
        Next key
    End If
    
End Function






Public Sub removeItem(dict As Scripting.Dictionary, key As Variant)
    With dict
        If .exists(key) Then Call .Remove(key)
    End With
End Sub

Public Sub removeNullValues(dict As Scripting.Dictionary)
    Dim key As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    If dict Is Nothing Then
    Else
        For Each key In dict.Keys
            If VBA.IsNull(dict.item(key)) Then
                Call dict.Remove(key)
            End If
        Next key
    End If
    
End Sub








'[MODIFYING CONTENT]

Public Function swapKeyWithValue(dict As Scripting.Dictionary, _
                                    Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Override) As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If Not dict Is Nothing Then
        Set swapKeyWithValue = F.Dictionaries.Create(dict.CompareMode = BinaryCompare)
        For Each key In dict.Keys
            Call F.Dictionaries.addItem(swapKeyWithValue, dict.item(key), key, duplicateBehaviour)
        Next key
    End If
    
End Function

Public Function changeKeys(dict As Scripting.Dictionary, functionName As String, _
                            Optional functionInvoker As Object, Optional functionParams As Scripting.Dictionary, _
                            Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Skip) As Scripting.Dictionary
    Dim key As Variant
    Dim newKey As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set changeKeys = F.Dictionaries.Create(dict.CompareMode = BinaryCompare)
    For Each key In dict.Keys
        If Not VBA.IsEmpty(key) Then
            Call F.Variables.assign(newKey, _
                    F.Reflection.modifyValue(dict.item(key), functionName, functionInvoker, functionParams))
            Call F.Dictionaries.addItem(changeKeys, newKey, dict.item(key), duplicateBehaviour)
        End If
    Next key
    
End Function











'[COMPARING DICTIONARIES]
Public Function compare(baseValue As Object, _
                                    comparedValue As Object, _
                                    Optional strictDataType As Boolean = False, _
                                    Optional ignoreCompareMode As Boolean = True, _
                                    Optional isCaseSensitive As Boolean = False) As Boolean
    Dim isBaseDictionary As Boolean:                isBaseDictionary = isDictionary(baseValue)
    Dim isComparedDictionary As Boolean:            isComparedDictionary = isDictionary(comparedValue)
    '------------------------------------------------------------------------------------------------------

    If Not checkIfAllItemsAreDictionaryOrNothing(baseValue, comparedValue) Then
        Call Exceptions.IllegalDataType.raise
    ElseIf Not (isBaseDictionary And isComparedDictionary) Then
        compare = (isBaseDictionary = isComparedDictionary)
    ElseIf baseValue.count <> comparedValue.count Then
        compare = False
    ElseIf Not ignoreCompareMode And baseValue.CompareMode <> comparedValue.CompareMode Then
        compare = False
    Else
        compare = compareDictionariesWithEqualNumberOfItems(baseValue, comparedValue, strictDataType, isCaseSensitive)
    End If

End Function

Private Function checkIfAllItemsAreDictionaryOrNothing(ParamArray values() As Variant) As Boolean
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    checkIfAllItemsAreDictionaryOrNothing = True
    For Each item In values
        If isDictionary(item) Or F.Variables.isNothing(item) Then
        Else
            checkIfAllItemsAreDictionaryOrNothing = False
            Exit For
        End If
    Next item
End Function

Private Function compareDictionariesWithEqualNumberOfItems(baseDict As Scripting.Dictionary, _
                                    comparedDict As Scripting.Dictionary, _
                                    Optional strictDataType As Boolean = False, _
                                    Optional isCaseSensitive As Boolean = False) As Boolean
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------

    compareDictionariesWithEqualNumberOfItems = True
    For Each key In baseDict.Keys
        If Not hasTheSameItem(comparedDict, key, baseDict.item(key), strictDataType, isCaseSensitive) Then
            compareDictionariesWithEqualNumberOfItems = False
            Exit For
        End If
    Next key

    'Counter-check if dictionaries have different comparing mode
    If compareDictionariesWithEqualNumberOfItems Then
        If baseDict.CompareMode <> comparedDict.CompareMode Then
            For Each key In comparedDict.Keys
                If Not hasTheSameItem(baseDict, key, comparedDict.item(key), strictDataType, isCaseSensitive) Then
                    compareDictionariesWithEqualNumberOfItems = False
                    Exit For
                End If
            Next key
        End If
    End If

End Function

Private Function hasTheSameItem(dict As Object, key As Variant, comparedItem As Variant, _
                                    Optional strictDataType As Boolean = False, _
                                    Optional isCaseSensitive As Boolean = False) As Boolean
    With dict
        If .exists(key) Then
            hasTheSameItem = F.Variables.equals(comparedItem, .item(key), strictDataType, isCaseSensitive)
        End If
    End With
End Function







'[CHECKING DICTIONARY CONTENT]
Public Function hasAllKeys(checkedDict As Scripting.Dictionary, comparedDict As Scripting.Dictionary) As Boolean
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If checkedDict Is Nothing Then
        hasAllKeys = False
    ElseIf comparedDict Is Nothing Then
        hasAllKeys = True
    Else
        hasAllKeys = True
        For Each key In comparedDict.Keys
            If Not checkedDict.exists(key) Then
                hasAllKeys = False
                Exit For
            End If
        Next key
    End If
    
End Function










'[OTHERS]
Public Function countByValue(dict As Scripting.Dictionary, searchValue As Variant, _
                             Optional functionName As String, _
                             Optional functionInvoker As Object) As Long
    Dim key As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    If isNonEmptyDictionary(dict) Then
        For Each key In dict.Keys
            Call F.Variables.assign(value, F.Reflection.modifyValue(dict.item(key), functionName, functionInvoker))
            If F.Variables.equals(searchValue, value) Then
                countByValue = countByValue + 1
            End If
        Next key
    End If
    
End Function


Public Function flatDictionary(dict As Scripting.Dictionary, Optional keyPrefix As String, Optional delimiter As String = ".") As Scripting.Dictionary
    Dim key As Variant
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    Dim resultDict As Scripting.Dictionary
    Dim subDict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set resultDict = F.Dictionaries.Create(False)
    
    If F.Dictionaries.isNonEmptyDictionary(dict) Then
        For Each key In dict.Keys
            Call F.Variables.assign(item, dict.item(key))
            If F.Dictionaries.isDictionary(item) Then
                Set subDict = F.cast.toDictionary(item)
                Call appendDictionary(resultDict, flatDictionary(subDict, keyPrefix & key & delimiter, delimiter))
            Else
                Call resultDict.Add(keyPrefix & key, item)
            End If
        Next key
    End If
    
    Set flatDictionary = resultDict
    
End Function


'#todo - check
'[old version]
'Public Function flatten(dict As Scripting.Dictionary, Optional delimiter As String = ".", _
'                                    Optional itemsPrefix As String = vbNullString) As Scripting.Dictionary
'    Dim varKey As Variant
'    Dim destinationKey As String
'    Dim item As Variant
'    Dim subDict As Scripting.Dictionary
'    '------------------------------------------------------------------------------------------------------
'    Set flatten = F.Dictionaries.Create(False)
'    For Each varKey In dict.keys
'        If VBA.Len(itemsPrefix) Then destinationKey = itemsPrefix & delimiter & varKey Else destinationKey = varKey
'        If isDictionary(dict.item(varKey)) Then
'            Set subDict = flatten(dict.item(varKey), delimiter, destinationKey)
'            Set flatten = join(flatten, subDict, True, False)
'        Else
'            Call flatten.Add(destinationKey, dict.item(varKey))
'        End If
'    Next varKey
'End Function









Public Function filter(dict As Scripting.Dictionary, _
                            Optional filterMethod As String = vbNullString, _
                            Optional filterMethodInvoker As Object, _
                            Optional filterMethodParams As Scripting.Dictionary, _
                            Optional value As Variant) As Scripting.Dictionary
    Dim key As Variant
    Dim filterValue As Variant
    Dim include As Variant
    '------------------------------------------------------------------------------------------------------
    
    If Not dict Is Nothing Then
        Set filter = F.Dictionaries.Create(dict.CompareMode = BinaryCompare)
        For Each key In dict.Keys
            If VBA.Len(filterMethod) Then
                include = F.Reflection.modifyValue(dict.item(key), filterMethod, filterMethodInvoker, filterMethodParams)
            Else
                include = F.Variables.equals(dict.item(key), value)
            End If
            
            If include Then
                Call filter.Add(key, dict.item(key))
            End If
        Next key
    Else
        Set filter = F.Dictionaries.Create(False)
    End If
    
End Function

















'
'
'Public Function join(baseDict As Scripting.Dictionary, _
'                appendedDict As Scripting.Dictionary, _
'                Optional ignoreDuplicates As Boolean = True, _
'                Optional cloneBaseDictionary As Boolean = True) As Scripting.Dictionary
'    Const METHOD_NAME As String = "join"
'    '------------------------------------------------------------------------------------------------------
'    Dim varKey As Variant
'    '------------------------------------------------------------------------------------------------------
'
'    If baseDict Is Nothing Then
'        If appendedDict Is Nothing Then
'            Set join = Create
'        Else
'            Set join = clone(appendedDict)
'        End If
'    Else
'        If cloneBaseDictionary Then Set join = clone(baseDict) Else Set join = baseDict
'        If Not appendedDict Is Nothing Then
'            With join
'                For Each varKey In appendedDict.keys
'                    If .exists(varKey) Then
'                        If Not ignoreDuplicates Then Call Exceptions.DuplicateKey.raise(varKey)
'                    Else
'                        Call .Add(varKey, appendedDict.item(varKey))
'                    End If
'                Next varKey
'            End With
'        End If
'    End If
'
'End Function
'
'
'
'
''[LOOKING FOR ITEMS]
'
'Public Function getKeysByItem(dict As Object, item As Variant, _
'                                                Optional strictDataType As Boolean = False, _
'                                                Optional isCaseSensitive As Boolean = False, _
'                                                Optional trimIfValueIsString As Boolean = True) As Variant()
'    Dim arr() As Variant
'    Dim varKey As Variant
'    '------------------------------------------------------------------------------------------------------
'
'    If dict Is Nothing Then
'
'    ElseIf Not isDictionary(dict) Then
'        Call Exceptions.IllegalDataType.raise
'    Else
'        For Each varKey In dict.keys
'            If F.Variables.equals(item, dict.item(varKey), strictDataType, isCaseSensitive, trimIfValueIsString) Then
'                Call F.arrays.addEntry(arr, varKey)
'            End If
'        Next varKey
'    End If
'
'    getKeysByItem = arr
'
'End Function
'
'
'
'
'
''***** WITHOUT UNIT TESTS *****
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'Public Function removeCommonItems(baseDict As Scripting.Dictionary, container As Variant, compareKeys As Boolean) As Scripting.Dictionary
'    Dim col As VBA.Collection
'    '------------------------------------------------------------------------------------------------------
'    Set col = F.Collections.Create(container)
'    If compareKeys Then
'        Call filterDictionaryBasedOnComparingKeysToGivenCollection(baseDict, col, False)
'    Else
'        Call filterDictionaryBasedOnComparingItemsToGivenCollection(baseDict, col, False)
'    End If
'    Set removeCommonItems = baseDict
'End Function
'
'Public Function leaveOnlyCommonItems(baseDict As Scripting.Dictionary, container As Variant, compareKeys As Boolean) As Scripting.Dictionary
'    Dim col As VBA.Collection
'    '------------------------------------------------------------------------------------------------------
'    Set col = F.Collections.Create(container)
'    If compareKeys Then
'        Call filterDictionaryBasedOnComparingKeysToGivenCollection(baseDict, col, True)
'    Else
'        Call filterDictionaryBasedOnComparingItemsToGivenCollection(baseDict, col, True)
'    End If
'    Set leaveOnlyCommonItems = baseDict
'End Function
'
'Private Sub filterDictionaryBasedOnComparingKeysToGivenCollection(dict As Scripting.Dictionary, _
'                            col As VBA.Collection, removeKeysMissingFromCollection As Boolean)
'    Dim varKey As Variant
'    '------------------------------------------------------------------------------------------------------
'    For Each varKey In dict.keys
'        If F.Collections.contains(col, varKey) <> removeKeysMissingFromCollection Then
'            Call dict.Remove(varKey)
'        End If
'    Next varKey
'End Sub
'
'Private Sub filterDictionaryBasedOnComparingItemsToGivenCollection(dict As Scripting.Dictionary, _
'                            col As VBA.Collection, removeItemsMissingFromCollection As Boolean)
'    Dim varKey As Variant
'    '------------------------------------------------------------------------------------------------------
'    For Each varKey In dict.keys
'        If F.Collections.contains(col, dict.item(varKey)) <> removeItemsMissingFromCollection Then
'            Call dict.Remove(varKey)
'        End If
'    Next varKey
'End Sub
'
'
'
'
'
'Public Function getValueByKeyRegex(dict As Scripting.Dictionary, regex As String) As Variant
'    Dim varKey As Variant
'    Dim objRegex As Object
'    '------------------------------------------------------------------------------------------------------
'
'    Set objRegex = F.regex.Create(regex)
'    For Each varKey In dict.keys
'        If objRegex.test(F.Strings.stringify(varKey)) Then
'            Call F.Variables.assign(getValueByKeyRegex, dict.item(varKey))
'            Exit For
'        End If
'    Next varKey
'
'End Function
'
'
'
''
''Private Function getFunctionResult(item As Variant, functionName As String, useInsideMethod As Boolean, _
''                                            Optional params As Scripting.Dictionary) As Variant
''    If useInsideMethod Then
''        On Error Resume Next
''        If Not params Is Nothing Then
''            getFunctionResult = VBA.CallByName(item, functionName, VbMethod, params)
''        Else
''            getFunctionResult = VBA.CallByName(item, functionName, VbMethod)
''        End If
''        On Error GoTo 0
''    Else
''        On Error Resume Next
''        If Not params Is Nothing Then
''            getFunctionResult = Application.run(functionName, item, params)
''        Else
''            getFunctionResult = Application.run(functionName, item)
''        End If
''        On Error GoTo 0
''    End If
''End Function
'
'
'
'
'
'Public Function getItemWithAnyGivenKey(dict As Object, keys As Variant, Optional returnNothingIfNotFound As Boolean = False) As Variant
'    Dim i As Long
'    Dim varItem As Variant
'    Dim itemFound As Boolean
'    Dim dictDict As Scripting.Dictionary
'    '------------------------------------------------------------------------------------------------------
'
'    If isDictionary(dict) Then
'        Set dictDict = dict
'        If F.arrays.isDefined(keys) Then
'            For i = LBound(keys) To UBound(keys)
'                If dictDict.exists(keys(i)) Then
'                    Call F.Variables.assign(getItemWithAnyGivenKey, dictDict.item(keys(i)))
'                    itemFound = True
'                    Exit For
'                End If
'            Next i
'        ElseIf F.Collections.isNonEmptyCollection(keys) Then
'            For Each varItem In keys
'                If dictDict.exists(varItem) Then
'                    Call F.Variables.assign(getItemWithAnyGivenKey, dictDict.item(varItem))
'                    itemFound = True
'                    Exit For
'                End If
'            Next varItem
'        Else
'
'        End If
'    End If
'
'    If Not itemFound And returnNothingIfNotFound Then getItemWithAnyGivenKey = Nothing
'
'End Function
'
'
'
'
'
'
