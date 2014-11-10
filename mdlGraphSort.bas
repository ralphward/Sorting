Attribute VB_Name = "mdlGraphSort"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Option Explicit

' Bubble Sort - steps through each element and swaps with the next one if required, keeps iterating until no swaps are left
Public Sub GrphBubbleSort(ByRef pvarArray As Variant, Optional ByRef MSChart As MSChart)
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
    
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray) - 1
    Do
        blnSwapped = False
        For i = iMin To iMax
            If pvarArray(i) > pvarArray(i + 1) Then
                varSwap = pvarArray(i)
                pvarArray(i) = pvarArray(i + 1)
                pvarArray(i + 1) = varSwap
                blnSwapped = True
                frmMain.Controls("MSChart").ChartData = pvarArray
                DoEvents
                Sleep 100
            End If
        Next
        iMax = iMax - 1
    Loop Until Not blnSwapped
End Sub

' Cocktail Sort - similar to bubble search but forward and backward through a collection
Public Sub GrphCocktailSort(ByRef pvarArray As Variant)
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
    
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray) - 1
    Do
        blnSwapped = False
        For i = iMin To iMax
            If pvarArray(i) > pvarArray(i + 1) Then
                varSwap = pvarArray(i)
                pvarArray(i) = pvarArray(i + 1)
                pvarArray(i + 1) = varSwap
                blnSwapped = True
                frmMain.Controls("MSChart").ChartData = pvarArray
                DoEvents
                Sleep 100
            End If
        Next
        iMax = iMax - 1
        If Not blnSwapped Then Exit Do
        For i = iMax To iMin Step -1
            If pvarArray(i) > pvarArray(i + 1) Then
                varSwap = pvarArray(i)
                pvarArray(i) = pvarArray(i + 1)
                pvarArray(i + 1) = varSwap
                blnSwapped = True
                frmMain.Controls("MSChart").ChartData = pvarArray
                DoEvents
                Sleep 100
            End If
        Next
        iMin = iMin + 1
    Loop Until Not blnSwapped
End Sub

' CombSort
Public Sub GrphCombSort(ByRef pvarArray As Variant)
    Const ShrinkFactor = 1.3
    Dim lngGap As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
    
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray)
    lngGap = iMax - iMin + 1
    Do
        If lngGap > 1 Then
            lngGap = Int(lngGap / ShrinkFactor)
            If lngGap = 10 Or lngGap = 9 Then lngGap = 11
        End If
        blnSwapped = False
        For i = iMin To iMax - lngGap
            If pvarArray(i) > pvarArray(i + lngGap) Then
                varSwap = pvarArray(i)
                pvarArray(i) = pvarArray(i + lngGap)
                pvarArray(i + lngGap) = varSwap
                blnSwapped = True
                frmMain.Controls("MSChart").ChartData = pvarArray
                DoEvents
                Sleep 100
            End If
        Next
    Loop Until lngGap = 1 And Not blnSwapped
End Sub

' Gnome Sort - swaps the first out of order elements and checks backwards to see if a swap should be made before going forward again
Public Sub GrphGnomeSort(ByRef pvarArray As Variant)
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    
    iMin = LBound(pvarArray) + 1
    iMax = UBound(pvarArray)
    i = iMin
    j = i + 1
    Do While i <= iMax
        If pvarArray(i) < pvarArray(i - 1) Then
            varSwap = pvarArray(i)
            pvarArray(i) = pvarArray(i - 1)
            pvarArray(i - 1) = varSwap
            frmMain.Controls("MSChart").ChartData = pvarArray
            DoEvents
            Sleep 100
            If i > iMin Then i = i - 1
         Else
            i = j
            j = j + 1
        End If
    Loop
End Sub

' Heap sort
Public Sub GrphHeapSort(ByRef pvarArray As Variant)
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray)
    For i = (iMax + iMin) \ 2 To iMin Step -1
        GrphHeap1 pvarArray, i, iMin, iMax
    Next i
    For i = iMax To iMin + 1 Step -1
        varSwap = pvarArray(i)
        pvarArray(i) = pvarArray(iMin)
        pvarArray(iMin) = varSwap
        GrphHeap1 pvarArray, iMin, iMin, i - 1
    Next i
End Sub

Private Sub GrphHeap1(ByRef pvarArray As Variant, ByVal i As Long, iMin As Long, iMax As Long)
    Dim lngLeaf As Long
    Dim varSwap As Variant
    
    Do
        lngLeaf = i + i - (iMin - 1)
        Select Case lngLeaf
            Case Is > iMax: Exit Do
            Case Is < iMax: If pvarArray(lngLeaf + 1) > pvarArray(lngLeaf) Then lngLeaf = lngLeaf + 1
        End Select
        If pvarArray(i) > pvarArray(lngLeaf) Then Exit Do
        varSwap = pvarArray(i)
        pvarArray(i) = pvarArray(lngLeaf)
        pvarArray(lngLeaf) = varSwap
        i = lngLeaf
    Loop
End Sub

' Insertion sort - every iteration of an insertion sort removes an element from the input data, inserting it at the correct position in the already sorted list,
' until no elements are left in the input
Public Sub GrphInsertionSort(ByRef pvarArray As Variant)
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    
    iMin = LBound(pvarArray) + 1
    iMax = UBound(pvarArray)
    For i = iMin To iMax
        varSwap = pvarArray(i)
        For j = i To iMin Step -1
            If varSwap < pvarArray(j - 1) Then
                pvarArray(j) = pvarArray(j - 1)
                frmMain.Controls("MSChart").ChartData = pvarArray
                DoEvents
                Sleep 100
            Else
                Exit For
            End If
        Next j
        pvarArray(j) = varSwap
        frmMain.Controls("MSChart").ChartData = pvarArray
        DoEvents
        Sleep 100
        
    Next i
End Sub

' Similar to Bubble sort as it uses nested loops - but moves elements very far initially - iteratively reducing the distance values are moved
Public Sub GrphJumpSort(ByRef pvarArray As Variant)
    Dim lngJump As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
    
    iMin = LBound(pvarArray)
    iMax = UBound(pvarArray)
    lngJump = iMax - iMin
    If lngJump < 2 Then lngJump = 2
    Do
        lngJump = lngJump \ 2
        Do
            blnSwapped = False
            For i = iMin To iMax - lngJump
                If pvarArray(i) > pvarArray(i + lngJump) Then
                    varSwap = pvarArray(i)
                    pvarArray(i) = pvarArray(i + lngJump)
                    pvarArray(i + lngJump) = varSwap
                    frmMain.Controls("MSChart").ChartData = pvarArray
                    DoEvents
                    Sleep 100
                    blnSwapped = True
                End If
            Next
        Loop Until Not blnSwapped
    Loop Until lngJump = 1
End Sub

' Merge sort continually divides the collection into smaller arrays until it has only 2 - then sorts those two and merges arrays back together.
' Memory overhead in this implementation that could be improved upon greatly
' Possible to use insertion sort once your array is between 10 - 50 but not implemented here
' Omit optional params when calling as they're used internally in the function during recursion
Public Sub GrphMergeSort(ByRef pvarArray As Variant, Optional pvarMirror As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngMid As Long
    Dim l As Long
    Dim R As Long
    Dim O As Long
    Dim varSwap As Variant

    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
        ReDim pvarMirror(plngLeft To plngRight)
    End If
    lngMid = plngRight - plngLeft
    Select Case lngMid
        Case 0
        Case 1
            If pvarArray(plngLeft) > pvarArray(plngRight) Then
                varSwap = pvarArray(plngLeft)
                pvarArray(plngLeft) = pvarArray(plngRight)
                pvarArray(plngRight) = varSwap
            End If
        Case Else
            lngMid = lngMid \ 2 + plngLeft
            GrphMergeSort pvarArray, pvarMirror, plngLeft, lngMid
            GrphMergeSort pvarArray, pvarMirror, lngMid + 1, plngRight
            ' Merge the resulting halves
            l = plngLeft ' start of first (left) half
            R = lngMid + 1 ' start of second (right) half
            O = plngLeft ' start of output (mirror array)
            Do
                If pvarArray(R) < pvarArray(l) Then
                    pvarMirror(O) = pvarArray(R)
                    R = R + 1
                    If R > plngRight Then
                        For l = l To lngMid
                            O = O + 1
                            pvarMirror(O) = pvarArray(l)
                        Next
                        Exit Do
                    End If
                Else
                    pvarMirror(O) = pvarArray(l)
                    l = l + 1
                    If l > lngMid Then
                        For R = R To plngRight
                            O = O + 1
                            pvarMirror(O) = pvarArray(R)
                        Next
                        Exit Do
                    End If
                End If
                O = O + 1
            Loop
            For O = plngLeft To plngRight
                pvarArray(O) = pvarMirror(O)
            Next
    End Select
End Sub

'QuickSort - To partition an array, a pivot element is first randomly selected, and then compared against every other element.
' All smaller elements are moved before the pivot, and all larger elements are moved after.
' The lesser and greater sublists are then recursively processed until the entire list is sorted. This can be done efficiently in linear time and in-place.
Public Sub GrphQuickSort(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant
    
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While pvarArray(lngFirst) < varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid < pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
            frmMain.Controls("MSChart").ChartData = pvarArray
            DoEvents
            Sleep 100
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then GrphQuickSort pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then GrphQuickSort pvarArray, lngFirst, plngRight
End Sub

