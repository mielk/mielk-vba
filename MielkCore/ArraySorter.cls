VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArraySorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Enum SortTypeEnum
    SortType_None = 0
    SortType_Bubble = 1
    SortType_Selection = 2
    SortType_Insertion = 3
    SortType_Merge = 4
    SortType_Quick = 5
    SortType_Radix = 6
End Enum

Private Const CLASS_NAME As String = "ArraySorter"
    '----------------------------------------------------------------------------------------------------------
Private pArray As Variant
Private pDimensions As Integer
Private pOrderAscending As Boolean
Private pSortColumn As Long
Private pIntegerValues As Boolean
Private pHasHeader As Boolean
    '----------------------------------------------------------------------------------------------------------




'[SETTERS]
Public Function setArray(arr As Variant) As ArraySorter
    Set setArray = Me
    If VBA.IsArray(arr) Then
        pArray = arr
        pDimensions = F.arrays.countDimensions(pArray)
    Else
        Call err.raise(1, "Array expected")
    End If
End Function


'[API]
Public Function sort(Optional ascending As Boolean = True, Optional sortColumn As Long = 1, _
                     Optional integerValues As Boolean = False, _
                     Optional hasHeader As Boolean = False) As Variant
    Dim sortType As SortTypeEnum
    '------------------------------------------------------------------------------------------------------
    
    pSortColumn = sortColumn
    pOrderAscending = ascending
    pIntegerValues = integerValues
    pHasHeader = hasHeader
    
    sortType = determineSortType
    
    If pDimensions = 1 Then
        Select Case sortType
            Case SortType_None:             sort = pArray
            Case SortType_Bubble:           sort = bubbleSort_1D(pArray)
            Case SortType_Merge:            sort = mergeSort_1D(pArray)
        End Select
    ElseIf pDimensions = 2 Then
        Select Case sortType
            Case SortType_None:             sort = pArray
            Case SortType_Bubble:           sort = bubbleSort_2D(pArray)
            Case SortType_Merge:            sort = mergeSort_2D(pArray)
        End Select
    End If
    
End Function





Private Function determineSortType() As SortTypeEnum
    Dim nextGreater As Long
    Dim nextLess As Long
    Dim nextEqual As Long
    Dim i As Long
    Dim itemsCounter As Long
    '------------------------------------------------------------------------------------------------------
    Dim bubbleSortComplexity As Double
    Dim mergeSortComplexity As Double
    '------------------------------------------------------------------------------------------------------
    
    itemsCounter = UBound(pArray, pDimensions) - LBound(pArray, pDimensions) + 1
    If pDimensions = 1 Then
        For i = LBound(pArray, 1) + VBA.IIf(pHasHeader, 1, 0) To UBound(pArray, 1) - 1
            If pArray(i) > pArray(i + 1) Then
                nextLess = nextLess + 1
            ElseIf pArray(i) = pArray(i + 1) Then
                nextEqual = nextEqual + 1
            Else
                nextGreater = nextGreater + 1
            End If
        Next i
    ElseIf pDimensions = 2 Then
        For i = LBound(pArray, 2) + VBA.IIf(pHasHeader, 1, 0) To UBound(pArray, 2) - 1
            If pArray(pSortColumn, i) > pArray(pSortColumn, i + 1) Then
                nextLess = nextLess + 1
            ElseIf pArray(pSortColumn, i) = pArray(pSortColumn, i + 1) Then
                nextEqual = nextEqual + 1
            Else
                nextGreater = nextGreater + 1
            End If
        Next i
    End If
       
    
    If pOrderAscending Then
        If nextLess = 0 Then
            determineSortType = SortType_None
        Else
            If nextLess > nextGreater Then
                pArray = F.arrays.invert(pArray, dimensions:=pDimensions)
                Call F.Variables.swapValues(nextGreater, nextLess)
            End If
            bubbleSortComplexity = calculateBubbleSortTimeComplexity(itemsCounter, nextGreater)
            mergeSortComplexity = calculateMergeSortTimeComplexity(itemsCounter)
            
            If mergeSortComplexity < bubbleSortComplexity Then
                determineSortType = SortType_Merge
            Else
                determineSortType = SortType_Bubble
            End If
            
        End If
    Else
        If nextGreater = 0 Then
            determineSortType = SortType_None
        Else
            If nextGreater > nextLess Then
                pArray = F.arrays.invert(pArray, dimensions:=pDimensions)
                Call F.Variables.swapValues(nextGreater, nextLess)
            End If
            bubbleSortComplexity = calculateBubbleSortTimeComplexity(itemsCounter, nextLess)
            mergeSortComplexity = calculateMergeSortTimeComplexity(itemsCounter)
            
            If mergeSortComplexity < bubbleSortComplexity Then
                determineSortType = SortType_Merge
            Else
                determineSortType = SortType_Bubble
            End If
            
        End If
    End If
    
    
End Function

Private Function calculateBubbleSortTimeComplexity(totalItems As Long, greaterValues As Long) As Double
    Dim serieFirstItem As Long: serieFirstItem = totalItems - greaterValues
    Dim serieLastItem As Long: serieLastItem = totalItems - 1
    '------------------------------------------------------------------------------------------------------
    calculateBubbleSortTimeComplexity = greaterValues * (serieFirstItem + serieLastItem) / 2
End Function

Private Function calculateMergeSortTimeComplexity(totalItems As Long) As Double
    calculateMergeSortTimeComplexity = Excel.Application.WorksheetFunction.RoundUp(Log(totalItems) / Log(2), 0) * totalItems
End Function




Private Function bubbleSort_1D(arr As Variant) As Variant 'arr As Variant, column As Long, Optional ascending As Boolean = True)
    Dim noSwaps As Boolean
    Dim i As Long
    Dim j As Long
    '------------------------------------------------------------------------------------------------------
    
    For i = UBound(arr) To LBound(arr) Step -1
        noSwaps = True
        For j = LBound(arr) To i - 1
            If (arr(j) > arr(j + 1)) = pOrderAscending Then
                Call F.arrays.swapItems(arr, j, j + 1, dimensions:=pDimensions)
                noSwaps = False
            End If
        Next j
        If noSwaps Then Exit For
    Next i
    bubbleSort_1D = arr
    
End Function

Private Function bubbleSort_2D(arr As Variant) As Variant 'arr As Variant, column As Long, Optional ascending As Boolean = True)
    Dim noSwaps As Boolean
    Dim i As Long
    Dim j As Long
    '------------------------------------------------------------------------------------------------------
    
    For i = UBound(arr, 2) To LBound(arr, 2) Step -1
        noSwaps = True
        For j = LBound(arr, 2) To i - 1
            If (arr(pSortColumn, j) > arr(pSortColumn, j + 1)) = pOrderAscending Then
                Call F.arrays.swapItems(arr, j, j + 1, dimensions:=pDimensions)
                noSwaps = False
            End If
        Next j
        If noSwaps Then Exit For
    Next i
    bubbleSort_2D = arr
    
End Function

Private Function insertionSort() As Variant 'arr As Variant, column As Long, Optional ascending As Boolean = True)

End Function

Private Function selectionSort() As Variant 'arr As Variant, column As Long, Optional ascending As Boolean = True)

End Function





Private Function mergeSort_1D(arr As Variant) As Variant ', column As Long, Optional ascending As Boolean = True) As Variant
    Dim mid As Long
    Dim left As Variant
    Dim right As Variant
    Dim counter As Long:        counter = UBound(arr, 1) - LBound(arr, 1) + 1
    '------------------------------------------------------------------------------------------------------
    
    If counter <= 1 Then
        mergeSort_1D = arr
    Else
        mid = VBA.Int(counter / 2)
        left = mergeSort_1D(F.arrays.slice(arr, LBound(arr, 1), mid, 1))
        right = mergeSort_1D(F.arrays.slice(arr, LBound(arr, 1) + mid, counter - mid, 1))
        mergeSort_1D = mergeArrays_1D(left, right)
    End If
    
End Function

Private Function mergeArrays_1D(arr1 As Variant, arr2 As Variant) As Variant ', sortColumn As Long, Optional ascending As Boolean = True) As Variant
    Dim result() As Variant
    Dim i As Long:                  i = LBound(arr1, 1)
    Dim j As Long:                  j = LBound(arr2, 1)
    Dim counter1 As Long:           counter1 = UBound(arr1, 1) - LBound(arr1, 1) + 1
    Dim counter2 As Long:           counter2 = UBound(arr2, 1) - LBound(arr2, 1) + 1
    Dim inserted As Long
    Dim k As Long
    '------------------------------------------------------------------------------------------------------
    
    ReDim result(1 To counter1 + counter2)
    Do While (i <= counter1 And j <= counter2)
        inserted = inserted + 1
        If ((arr2(j) >= arr1(i)) = pOrderAscending) Then
            Call F.Variables.assign(result(inserted), arr1(i))
            i = i + 1
        Else
            Call F.Variables.assign(result(inserted), arr2(j))
            j = j + 1
        End If
    Loop
    
    Do While (i <= counter1)
        inserted = inserted + 1
        Call F.Variables.assign(result(inserted), arr1(i))
        i = i + 1
    Loop
    
    Do While (j <= counter2)
        inserted = inserted + 1
        Call F.Variables.assign(result(inserted), arr2(j))
        j = j + 1
    Loop
    
    mergeArrays_1D = result
    
End Function

Private Function mergeSort_2D(arr As Variant) As Variant ', column As Long, Optional ascending As Boolean = True) As Variant
    Dim mid As Long
    Dim left As Variant
    Dim right As Variant
    Dim counter As Long:        counter = UBound(arr, 2) - LBound(arr, 2) + 1
    '------------------------------------------------------------------------------------------------------
    
    If counter <= 1 Then
        mergeSort_2D = arr
    Else
        mid = VBA.Int(counter / 2)
        left = mergeSort_2D(F.arrays.slice(arr, LBound(arr, 2), mid, 2))
        right = mergeSort_2D(F.arrays.slice(arr, LBound(arr, 2) + mid, counter - mid, 2))
        mergeSort_2D = mergeArrays_2D(left, right)
    End If
    
End Function

Private Function mergeArrays_2D(arr1 As Variant, arr2 As Variant) As Variant ', sortColumn As Long, Optional ascending As Boolean = True) As Variant
    Dim result() As Variant
    Dim i As Long:                  i = LBound(arr1, 2)
    Dim j As Long:                  j = LBound(arr2, 2)
    Dim counter1 As Long:           counter1 = UBound(arr1, 2) - LBound(arr1, 2) + 1
    Dim counter2 As Long:           counter2 = UBound(arr2, 2) - LBound(arr2, 2) + 1
    Dim inserted As Long
    Dim k As Long
    '------------------------------------------------------------------------------------------------------
    
    ReDim result(LBound(arr2, 1) To UBound(arr2, 1), 1 To counter1 + counter2)
    Do While (i <= counter1 And j <= counter2)
        inserted = inserted + 1
        If ((arr2(pSortColumn, j) >= arr1(pSortColumn, i)) = pOrderAscending) Then
            For k = LBound(arr2, 1) To UBound(arr2, 1)
                Call F.Variables.assign(result(k, inserted), arr1(k, i))
            Next k
            i = i + 1
        Else
            For k = LBound(arr2, 1) To UBound(arr2, 1)
                Call F.Variables.assign(result(k, inserted), arr2(k, j))
            Next k
            j = j + 1
        End If
    Loop
    
    Do While (i <= counter1)
        inserted = inserted + 1
        For k = LBound(arr2, 1) To UBound(arr2, 1)
            Call F.Variables.assign(result(k, inserted), arr1(k, i))
        Next k
        i = i + 1
    Loop
    
    Do While (j <= counter2)
        inserted = inserted + 1
        For k = LBound(arr2, 1) To UBound(arr2, 1)
            Call F.Variables.assign(result(k, inserted), arr2(k, j))
        Next k
        j = j + 1
    Loop
    
    mergeArrays_2D = result
    
End Function

Private Function quickSort() As Variant 'arr As Variant, column As Long, Optional ascending As Boolean = True)

End Function

Private Function radixSort() As Variant 'arr As Variant, column As Long, Optional ascending As Boolean = True)

End Function
