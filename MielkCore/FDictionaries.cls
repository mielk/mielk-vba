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
Public Function Create(Optional isCaseSensitive As Boolean = False) As Object
    Set Create = New Scripting.Dictionary
    Create.CompareMode = VBA.IIf(isCaseSensitive, DictCompareMode_Binary, DictCompareMode_Text)
End Function

Public Function createWithItems(isCaseSensitive As Boolean, _
                                ParamArray keyValuePairs() As Variant) As Scripting.Dictionary
    Dim varItem As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set createWithItems = Create(isCaseSensitive)
    For Each varItem In keyValuePairs
        Call f.dictionaries.addItem(createWithItems, varItem(1), varItem(2), duplicateBehaviour_ThrowError)
    Next varItem
    
End Function

Public Function createSingleEntryDict(key As Variant, value As Variant) As Scripting.Dictionary
    Set createSingleEntryDict = f.dictionaries.Create(False)
    Call createSingleEntryDict.Add(key, value)
End Function

Public Function createFrom(container As Variant, _
                            Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Skip, _
                            Optional keyFunction As String = vbNullString, Optional keyFunctionInvoker As Object) _
                            As Scripting.Dictionary

    If VBA.IsMissing(container) Then
        Set createFrom = Create(False)
    ElseIf VBA.IsEmpty(container) Then
        Set createFrom = Create(False)
    ElseIf f.Variables.isString(container) Then
        Set createFrom = fromString(VBA.CStr(container))
    ElseIf VBA.IsArray(container) Then
        Set createFrom = f.arrays.toDictionary(container, duplicateBehaviour:=duplicateBehaviour, _
                                                keyFunction:=keyFunction, keyFunctionInvoker:=keyFunctionInvoker)
    ElseIf Not f.Variables.isContainer(container) Then
        Set createFrom = f.dictionaries.createSingleEntryDict(container, container)
    ElseIf f.dictionaries.isDictionary(container) Then
        Set createFrom = container
    ElseIf f.Collections.isCollection(container) Then
        Set createFrom = f.Collections.toDictionary(f.cast.toCollection(container), duplicateBehaviour, keyFunction, keyFunctionInvoker)
    End If

End Function

Public Function fromParamArray(ParamArray values() As Variant) As Scripting.Dictionary
    Dim item As Variant
    '------------------------------------------------------------------------------------------------------
    Set fromParamArray = Create(False)
    For Each item In values
        Call addItem(fromParamArray, item, item, duplicateBehaviour_Override)
    Next item
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
    
    Set fromString = f.dictionaries.Create(False)
    
    regexPattern = f.strings.format(REGEX_PATTERN, keyItemSeparator, entriesSeparator)
    Set entries = f.regex.getMatchArraysCollection(text, regexPattern)
    For Each arr In entries
        key = arr(LBound(arr))
        value = arr(UBound(arr))
        If VBA.isNumeric(value) Then value = value * 1
        Call f.dictionaries.addItem(fromString, key, value, behaviourOnDuplicate)
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

Public Function getObject(dict As Scripting.Dictionary, key As Variant) As Object
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
            Call f.Variables.assign(getFirstEntry, key)
        ElseIf partType = DictPart_ValueOnly Then
            Call f.Variables.assign(getFirstEntry, dict.item(key))
        Else
            getFirstEntry = KeyValue(key, dict.item(key))
        End If
        Exit For
    Next key
    
End Function


Public Sub assignValueFromDictToVariable(ByVal dict As Scripting.Dictionary, ByRef variable As Variant, tags As String, _
                                            Optional conversionFunction As String)
    Const TAGS_SEPARATOR As String = ";"
    '------------------------------------------------------------------------------------------------------
    Dim arrTags() As String
    Dim value As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------

    arrTags = VBA.Split(tags, TAGS_SEPARATOR)

    If f.arrays.isDefined(arrTags) Then
        For i = LBound(arrTags) To UBound(arrTags)
            Call f.Variables.assign(value, getValueForCombinedKey(dict, VBA.trim$(arrTags(i)), conversionFunction))
            If Not VBA.IsEmpty(value) Then
                Call f.Variables.assign(variable, value)
                Exit For
            End If
        Next i
    End If

End Sub

Public Function getObjectForCombinedKey(ByVal dict As Scripting.Dictionary, tag As String, Optional conversionFunction As String) As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Call f.Variables.assign(value, getValueForCombinedKey(dict, tag, conversionFunction))
    If VBA.IsObject(value) Then
        Set getObjectForCombinedKey = value
    Else
        Set getObjectForCombinedKey = Nothing
    End If
    
End Function

Public Function getItemForCombinedKey(ByVal dict As Scripting.Dictionary, tag As String, Optional conversionFunction As String) As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    Call f.Variables.assign(value, getValueForCombinedKey(dict, tag, conversionFunction))
    If VBA.IsObject(value) Then
        getItemForCombinedKey = Empty
    Else
        getItemForCombinedKey = value
    End If
    
End Function

Private Function getValueForCombinedKey(ByVal dict As Scripting.Dictionary, tag As String, Optional conversionFunction As String) As Variant
    Dim arrTags As Variant
    Dim i As Long
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    arrTags = VBA.Split(tag, ".")
    
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        If f.arrays.isDefined(arrTags) Then
            For i = LBound(arrTags) To UBound(arrTags)
                With dict
                    If .exists(arrTags(i)) Then
                        If i = UBound(arrTags) Then
                            If VBA.Len(conversionFunction) Then
                                On Error GoTo ConversionException
                                Call f.Variables.assign(getValueForCombinedKey, Application.run(conversionFunction, .item(arrTags(i))))
                                If Not DEV_MODE Then On Error GoTo ErrHandler
                            Else
                                Call f.Variables.assign(getValueForCombinedKey, .item(arrTags(i)))
                            End If
                        Else
                            Set dict = .item(arrTags(i))
                        End If
                    Else
                        Exit For
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
    
ErrHandler:
    
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
    ElseIf f.dictionaries.isDictionary(value) Then
        Call clone_deepCloning(value, dict, key)
    ElseIf f.Collections.isCollection(value) Then
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
    If f.dictionaries.isNonEmptyDictionary(sourceContainer) Then
        Set dict = f.dictionaries.Create(False)
        Call destination.Add(key, dict)
        For Each varKey In sourceContainer.Keys
            Call clone_assignValue(dict, varKey, sourceContainer.item(varKey), True)
        Next varKey
    ElseIf f.Collections.isNonEmptyCollection(sourceContainer) Then
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
        Set sortByKey = f.dictionaries.Create(dict.CompareMode = TextCompare)
        If dict.count Then
            ReDim arr(1 To 2, 1 To dict.count)
            For Each key In dict.Keys
                i = i + 1
                Call f.Variables.assign(arr(1, i), key)
                Call f.Variables.assign(arr(2, i), dict.item(key))
            Next key
            
            arr = f.arrays.sort(arr, ascendingOrder, 1, False)
            
            For i = LBound(arr, 2) To UBound(arr, 2)
                Call sortByKey.Add(arr(1, i), arr(2, i))
            Next i
            
        End If
    End If
    
End Function

Public Function sortByValues(dict As Scripting.Dictionary, Optional ascendingOrder As Boolean = True) As Scripting.Dictionary
    Stop
End Function





'[CONVERTING COLLECTION INTO OTHER DATA TYPES]

Public Function toString(dict As Object, _
                                        Optional part As DictPartEnum = DictPart_KeyAndValue, _
                                        Optional stringifyMode As StringifyModeEnum = StringifyMode_Normal, _
                                        Optional itemsDelimiter As String = ", ", _
                                        Optional keyItemDelimiter As String = ":", _
                                        Optional openWrapper As String = "", _
                                        Optional closeWrapper As String = "") As String
    Dim entries() As String
    '------------------------------------------------------------------------------------------------------

    If dict Is Nothing Then
        toString = f.strings.stringify(Nothing, stringifyMode)
    ElseIf Not isDictionary(dict) Then
        Call Exceptions.IllegalDataType.raise
    Else
        entries = toString_getEntriesArray(dict, part, stringifyMode, keyItemDelimiter, openWrapper, closeWrapper)
        If stringifyMode = StringifyMode_Xml Then
            toString = f.strings.format(XML_DICTIONARY, VBA.join(entries, vbNullString))
        Else
            toString = VBA.join(entries, itemsDelimiter)
        End If
    End If

End Function

Private Function toString_getEntriesArray(dict As Object, part As DictPartEnum, _
                                        stringifyMode As StringifyModeEnum, _
                                        ByVal keyItemDelimiter As String, _
                                        ByVal openWrapper As String, _
                                        ByVal closeWrapper As String) As String()
    Dim varKey As Variant
    Dim values() As String
    Dim Counter As Long
    '------------------------------------------------------------------------------------------------------

    If dict.count > 0 Then
    
        If stringifyMode = StringifyMode_Xml Then
            openWrapper = vbNullString
            closeWrapper = vbNullString
        End If

        ReDim values(1 To dict.count)
        For Each varKey In dict.Keys
            Counter = Counter + 1
            values(Counter) = openWrapper & _
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
        getItemText = openTag & f.strings.stringify(varKey, stringifyMode) & closeTag
    ElseIf part = DictPart_ValueOnly Then
        Call f.Variables.assign(value, dict.item(varKey))
        getItemText = openTag & f.strings.stringify(value, stringifyMode) & closeTag
    ElseIf part = DictPart_KeyAndValue Then
        Call f.Variables.assign(value, dict.item(varKey))
        getItemText = getItemText_KeyValue(varKey, value, stringifyMode, keyItemDelimiter)
    End If

End Function

Private Function getItemText_KeyValue(key As Variant, value As Variant, _
                                                stringifyMode As StringifyModeEnum, _
                                                ByVal keyItemDelimiter As String) As String
    Dim keyText As String
    '------------------------------------------------------------------------------------------------------

    If stringifyMode = StringifyMode_Xml Then
        keyText = f.strings.stringify(key, StringifyMode_Normal)
        getItemText_KeyValue = "<" & keyText & ">" & _
                                f.strings.stringify(value, stringifyMode) & _
                                "</" & keyText & ">"
    Else
        getItemText_KeyValue = f.strings.stringify(key, StringifyMode_Normal) & _
                                keyItemDelimiter & _
                                f.strings.stringify(value, stringifyMode)
    End If

End Function



Public Function toArray(dict As Scripting.Dictionary, Optional dictPart As DictPartEnum = DictPart_KeyAndValue) As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------

    If dict Is Nothing Then
        ' return Empty
    ElseIf dict.count = 0 Then
        toArray = arr
    Else
        If dictPart = DictPart_KeyAndValue Then
            toArray = f.arrays.to2D(dict)
        ElseIf dictPart = DictPart_KeyOnly Then
            toArray = dict.Keys
        ElseIf dictPart = DictPart_ValueOnly Then
            toArray = dict.items
        End If
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
                    
'    If dict Is Nothing Then
'        Call Exceptions.NullNotAllowed.raiseWithCustomMessage("Dictionary cannot be Nothing")
'    End If
    
    If Not dict Is Nothing Then
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
    End If
    
End Sub

Public Sub appendItemsFromContainer(dict As Scripting.Dictionary, container As Variant, _
                    Optional behaviourIfDuplicate As DuplicateBehaviourEnum = duplicateBehaviour_Override)
    
    If Not f.Variables.isContainer(container) Then
        Call Exceptions.IllegalDataType.raiseWithCustomMessage("Given parameters [container] is not container")
    End If
    
    If Not dict Is Nothing Then
        If f.dictionaries.isDictionary(container) Then
            Call appendDictionary(dict, f.cast.toDictionary(container), behaviourIfDuplicate)
        ElseIf f.Collections.isNonEmptyCollection(container) Then
            Call appendCollection(dict, f.cast.toCollection(container), behaviourIfDuplicate)
        ElseIf f.arrays.isDefined(container) Then
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
    
    If f.Collections.isNonEmptyCollection(col) Then
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
    
    dimensions = f.arrays.countDimensions(arr)
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
    Dim subdict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If Not masterDict Is Nothing Then
        Set merge = f.dictionaries.clone(masterDict)
    Else
        Set merge = f.dictionaries.Create(False)
    End If
    
    If Not appendedDict Is Nothing Then
        For Each key In appendedDict
            If Not merge.exists(key) Then
                Call merge.Add(key, appendedDict.item(key))
            ElseIf f.dictionaries.isDictionary(merge.item(key)) And f.dictionaries.isDictionary(appendedDict.item(key)) Then
                Set subdict = merge(merge.item(key), appendedDict.item(key), deepMerge, duplicateBehaviour)
                Set merge.item(key) = subdict
            Else
                Call f.dictionaries.addItem(merge, key, appendedDict.item(key), duplicateBehaviour)
            End If
        Next key
    End If
    
End Function






Public Sub removeItem(dict As Scripting.Dictionary, key As Variant)
    With dict
        If .exists(key) Then Call .Remove(key)
    End With
End Sub

Public Sub removeNullAndEmptyValues(dict As Scripting.Dictionary, Optional removeNull As Boolean = True, _
                                                                  Optional removeEmpty As Boolean = True)
    Dim key As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    
    If dict Is Nothing Then
    Else
        For Each key In dict.Keys
            If removeNull And VBA.IsNull(dict.item(key)) Then
                Call dict.Remove(key)
            ElseIf removeEmpty And f.Variables.isUndefined(dict.item(key), True) Then
                Call dict.Remove(key)
            End If
        Next key
    End If

End Sub





'[MODIFYING CONTENT]

Public Function flip(dict As Scripting.Dictionary, _
                                    Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Override) As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If Not dict Is Nothing Then
        Set flip = f.dictionaries.Create(dict.CompareMode = BinaryCompare)
        For Each key In dict.Keys
            Call f.dictionaries.addItem(flip, dict.item(key), key, duplicateBehaviour)
        Next key
    End If
    
End Function

Public Function changeKeys(dict As Scripting.Dictionary, functionName As String, _
                            Optional functionInvoker As Object, Optional functionParams As Scripting.Dictionary, _
                            Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Skip, _
                            Optional useValueAsBaseForConvertion As Boolean = True) As Scripting.Dictionary
    Dim key As Variant
    Dim newKey As Variant
    Dim valueToBeConverted As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set changeKeys = f.dictionaries.Create(dict.CompareMode = BinaryCompare)
    For Each key In dict.Keys
        If Not VBA.IsEmpty(key) Then
            Call f.Variables.assign(valueToBeConverted, VBA.IIf(useValueAsBaseForConvertion, dict.item(key), key))
            Call f.Variables.assign(newKey, _
                    f.Reflection.modifyValue(valueToBeConverted, functionName, functionInvoker, functionParams))
            Call f.dictionaries.addItem(changeKeys, newKey, dict.item(key), duplicateBehaviour)
        End If
    Next key
    
End Function

Public Function repeatKeyAsValue(dict As Scripting.Dictionary) As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    Set repeatKeyAsValue = f.dictionaries.Create(False)
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        For Each key In dict.Keys
            Call f.dictionaries.addItem(repeatKeyAsValue, key, key)
        Next key
    End If
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
        If isDictionary(item) Or f.Variables.isNothing(item) Then
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
            hasTheSameItem = f.Variables.equals(comparedItem, .item(key), strictDataType, isCaseSensitive)
        End If
    End With
End Function

Public Function compareKeys(ByVal baseDict As Scripting.Dictionary, _
                            ByVal comparedDict As Scripting.Dictionary) As Scripting.Dictionary
    Dim key As Variant
    Dim dictCommon As Scripting.Dictionary
    Dim dictOnlyBase As Scripting.Dictionary
    Dim dictOnlyCompared As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set compareKeys = f.dictionaries.Create(False)
    
    Set baseDict = f.dictionaries.clone(baseDict)
    Set comparedDict = f.dictionaries.clone(comparedDict)
    
    If baseDict Is Nothing Then
        '[Only compared dictionary is valid]
        If Not comparedDict Is Nothing Then
            Call compareKeys.Add(ComparisonResult_OnlyInCompared, comparedDict)
        End If
    ElseIf comparedDict Is Nothing Then
        '[Only base range is valid]
        If Not baseDict Is Nothing Then
            Call compareKeys.Add(ComparisonResult_OnlyInBase, baseDict)
        End If
    Else
        Set dictCommon = f.dictionaries.Create(False)
        Set dictOnlyBase = f.dictionaries.Create(False)
        Set dictOnlyCompared = f.dictionaries.Create(False)
        
        '[Iterate through base dictionary]
        For Each key In baseDict.Keys
            If comparedDict.exists(key) Then
                Call dictCommon.Add(key, 0)
                Call comparedDict.Remove(key)
            Else
                Call dictOnlyBase.Add(key, 0)
            End If
        Next key
        
        '[Iterate through compare dictionary]
        For Each key In comparedDict.Keys
            Call dictOnlyCompared.Add(key, 0)
        Next key
        
    End If
    
    
    Set compareKeys = f.dictionaries.Create(False)
    With compareKeys
        Call .Add(ComparisonResult_Common, dictCommon)
        Call .Add(ComparisonResult_OnlyInBase, dictOnlyBase)
        Call .Add(ComparisonResult_OnlyInCompared, dictOnlyCompared)
    End With
    
End Function





'[CHECKING DICTIONARY CONTENT]
Public Function hasKey(dict As Scripting.Dictionary, key As Variant) As Boolean
    If Not dict Is Nothing Then
        hasKey = dict.exists(key)
    End If
End Function

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

Public Function hasCombinedKey(dict As Scripting.Dictionary, key As String) As Boolean
    Dim keyParts As Variant
    Dim subdict As Scripting.Dictionary
    Dim i As Long
    Dim part As String
    '------------------------------------------------------------------------------------------------------
    
    If Not dict Is Nothing Then
        keyParts = VBA.Split(key, PROPERTIES_SEPARATOR)
        Set subdict = dict
        
        For i = LBound(keyParts) To UBound(keyParts)
            part = keyParts(i)
            If Not subdict.exists(part) Then
                hasCombinedKey = False
                Exit For
            ElseIf i = UBound(keyParts) Then
                hasCombinedKey = True
                Exit For
            ElseIf TypeOf subdict.item(part) Is Scripting.Dictionary Then
                Set subdict = subdict.item(part)
            Else
                hasCombinedKey = False
                Exit For
            End If
        Next i
    
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
            Call f.Variables.assign(value, f.Reflection.modifyValue(dict.item(key), functionName, functionInvoker))
            If f.Variables.equals(searchValue, value) Then
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
    Dim subdict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set resultDict = f.dictionaries.Create(False)
    
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        For Each key In dict.Keys
            Call f.Variables.assign(item, dict.item(key))
            If f.dictionaries.isDictionary(item) Then
                Set subdict = f.cast.toDictionary(item)
                Call appendDictionary(resultDict, flatDictionary(subdict, keyPrefix & key & delimiter, delimiter))
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
        Set filter = f.dictionaries.Create(dict.CompareMode = BinaryCompare)
        For Each key In dict.Keys
            If VBA.Len(filterMethod) Then
                include = f.Reflection.modifyValue(dict.item(key), filterMethod, filterMethodInvoker, filterMethodParams)
            Else
                include = f.Variables.equals(dict.item(key), value)
            End If
            
            If include Then
                Call filter.Add(key, dict.item(key))
            End If
        Next key
    Else
        Set filter = f.dictionaries.Create(False)
    End If
    
End Function





Public Function getDictionaryDifference(baseDict As Scripting.Dictionary, comparedDict As Scripting.Dictionary) As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set getDictionaryDifference = f.dictionaries.Create(False)
    
    If comparedDict Is Nothing Then
        Set getDictionaryDifference = f.dictionaries.clone(baseDict)
    Else
        For Each key In baseDict.Keys
            If Not comparedDict.exists(key) Then
                Call getDictionaryDifference.Add(key, baseDict.item(key))
            ElseIf Not f.Variables.equals(baseDict.item(key), comparedDict.item(key), True) Then
                Call getDictionaryDifference.Add(key, baseDict.item(key))
            End If
        Next key
    End If
    
End Function


Public Function getKeyByItem(dict As Scripting.Dictionary, item As Variant) As Variant
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        For Each key In dict.Keys
            If f.Variables.equals(dict.item(key), item) Then
                Call f.Variables.assign(getKeyByItem, key)
                Exit For
            End If
        Next key
    End If
    
End Function


Public Function join(ParamArray dictionaries() As Variant) As Scripting.Dictionary
    Dim item As Variant
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set join = f.dictionaries.Create(False)
    For Each item In dictionaries
        If f.dictionaries.isNonEmptyDictionary(item) Then
            Set dict = f.cast.toDictionary(item)
            For Each key In dict.Keys
                Call f.dictionaries.addItem(join, key, dict.item(key))
            Next key
        End If
    Next item
    
End Function


Public Function replaceKey(dict As Scripting.Dictionary, oldKey As Variant, newKey As Variant) As Scripting.Dictionary
    Dim key As Variant
    Dim found As Boolean
    '------------------------------------------------------------------------------------------------------
    
    Set replaceKey = f.dictionaries.Create
    
    If isNonEmptyDictionary(dict) Then
        replaceKey.CompareMode = dict.CompareMode
        For Each key In dict.Keys
            If found Then
                Call f.dictionaries.addItem(replaceKey, key, dict.item(key))
            ElseIf f.Variables.equals(key, oldKey, True) Then
                Call f.dictionaries.addItem(replaceKey, newKey, dict.item(oldKey))
                found = True
            Else
                Call replaceKey.Add(key, dict.item(key))
            End If
        Next key
    End If
    
End Function


Public Function map(dict As Scripting.Dictionary, modifyMethod As String, _
                        Optional modifyMethodInvoker As Object, _
                        Optional modifyMethodParams As Scripting.Dictionary) As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    If Not dict Is Nothing Then
        Set map = f.dictionaries.Create(dict.CompareMode = BinaryCompare)
        For Each key In dict.Keys
            Call map.Add(key, f.Reflection.modifyValue(dict.item(key), _
                                            modifyMethod, modifyMethodInvoker, modifyMethodParams))
        Next key
    End If
End Function

Public Function getKeysForValue(dict As Scripting.Dictionary, value As Variant) As VBA.Collection
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    Set getKeysForValue = New VBA.Collection
    
    If isNonEmptyDictionary(dict) Then
        For Each key In dict.Keys
            If f.Variables.equals(dict.item(key), value) Then
                Call getKeysForValue.Add(key)
            End If
        Next key
    End If
    
End Function

Public Function hasValue(dict As Scripting.Dictionary, value As Variant) As Boolean
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------

    If isNonEmptyDictionary(dict) Then
        For Each key In dict.Keys
            If f.Variables.equals(dict.item(key), value) Then
                hasValue = True
                Exit For
            End If
        Next key
    End If

End Function


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

Public Function removeDuplicates(baseDict As Scripting.Dictionary, container As Variant, _
                          Optional compareByValue As Boolean = False) As Scripting.Dictionary
    Set removeDuplicates = removeItemsFromDictionary(baseDict, container, True, compareByValue)
End Function

Public Function removeNonDuplicates(baseDict As Scripting.Dictionary, container As Variant, _
                          Optional compareByValue As Boolean = False) As Scripting.Dictionary
    Set removeNonDuplicates = removeItemsFromDictionary(baseDict, container, False, compareByValue)
End Function

Private Function removeItemsFromDictionary( _
                        baseDict As Scripting.Dictionary, _
                        container As Variant, _
                        removeDuplicates As Boolean, _
                        Optional compareByValue As Boolean = False) As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    Dim comparedValue As Variant
    '------------------------------------------------------------------------------------------------------
    
    If Not baseDict Is Nothing Then
        Set removeItemsFromDictionary = f.dictionaries.Create(baseDict.CompareMode = BinaryCompare)
        Set dict = f.dictionaries.swapKeyWithValue( _
                        f.dictionaries.createFrom(container, duplicateBehaviour_Skip))
        
        For Each key In baseDict.Keys
            Call f.Variables.assign(comparedValue, VBA.IIf(compareByValue, baseDict.item(key), key))
            If dict.exists(comparedValue) <> removeDuplicates Then
                Call removeItemsFromDictionary.Add(key, baseDict.item(key))
            End If
        Next key
    
    End If
    
End Function



Public Function getValueByKeyRegex(dict As Scripting.Dictionary, regex As String) As Variant
    Dim varKey As Variant
    Dim objRegex As Object
    '------------------------------------------------------------------------------------------------------

    Set objRegex = f.regex.Create(regex)
    For Each varKey In dict.Keys
        If objRegex.test(f.strings.stringify(varKey)) Then
            Call f.Variables.assign(getValueByKeyRegex, dict.item(varKey))
            Exit For
        End If
    Next varKey

End Function


Public Sub removeKey(dict As Scripting.Dictionary, key As Variant)
    With dict
        If .exists(key) Then Call .Remove(key)
    End With
End Sub



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

Public Function swapKeyWithValue(dict As Scripting.Dictionary, _
                        Optional duplicateBehaviour As DuplicateBehaviourEnum = duplicateBehaviour_Skip) As Scripting.Dictionary
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    If Not dict Is Nothing Then
        Set swapKeyWithValue = f.dictionaries.Create(dict.CompareMode = BinaryCompare)
        For Each key In dict.Keys
            Call f.dictionaries.addItem(swapKeyWithValue, dict.item(key), key, duplicateBehaviour)
        Next key
    End If
End Function


Public Function getValuesSum(dict As Scripting.Dictionary) As Long
    Dim key As Variant
    Dim value As Variant
    '------------------------------------------------------------------------------------------------------
    If Not dict Is Nothing Then
        For Each key In dict.Keys
            value = dict.item(key)
            If f.Math.isNumeric(value) Then
                getValuesSum = getValuesSum + f.Math.toNumber(value)
            End If
        Next key
    End If
End Function


Public Sub addItemForCombinedKey(dict As Scripting.Dictionary, tag As String, value As Variant)
    Dim subdict As Scripting.Dictionary
    Dim arrTags As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    arrTags = VBA.Split(tag, ".")
    
    If f.dictionaries.isNonEmptyDictionary(dict) Then
        Set subdict = dict
        If f.arrays.isDefined(arrTags) Then
            For i = LBound(arrTags) To UBound(arrTags)
                If i = UBound(arrTags) Then
                    Call f.dictionaries.addItem(subdict, arrTags(i), value, duplicateBehaviour_Override)
                ElseIf subdict.exists(arrTags(i)) Then
                    Set subdict = subdict.item(arrTags(i))
                Else
                    Exit For
                End If
            Next i
        End If
    End If
    
'==========================================================================================================
ExitPoint:
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    
End Sub
