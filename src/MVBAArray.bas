Attribute VB_Name = "MVBAArray"
' ==========================================================================
' Module      : MVBAArray
' Type        : Module
' Description : Procedures for working with arrays
' --------------------------------------------------------------------------
' Procedures  : ArrayToString           String
'               BubbleSortArray
'               CollectionToArray       Variant
'               CombineArrays           Variant
'               GetArrayValue           Variant
'               GetElementIndex         Long
'               GetElementIndexX        Long
'               InsertElement           Boolean
'               IsAllocated             Boolean
'               IsEmptyArray            Boolean
'               IsInArray               Boolean
'               ListArray
'               NumberOfDimensions      Long
'               NumberOfElements        Long
'               OneDimToTwo             Variant
'               QuickSortArray
'               RemoveElement
'               StringToArray           Variant
'               UniqueItemsInArray      Variant
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE As String = "MVBAArray"

Public Function ArrayToString(ByRef Arr As Variant, _
                     Optional ByVal Delimiter As String = " ") As String
' ==========================================================================
' Description : Convert a 1-dimensional array to a string
'
' Parameters  : Arr         The array to convert
'               Delimiter   The delimiter to place between elements
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "ArrayToString"

    Dim sRtn    As String


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    sRtn = Join(Arr, Delimiter)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ArrayToString = sRtn

    Call Trace(tlMaximum, msMODULE, sPROC, sRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Sub BubbleSortArray(ByRef Arr As Variant, _
                           Optional ByVal DimIdx As Long, _
                           Optional ByVal SortOrder As XlSortOrder = xlAscending, _
                           Optional ByVal SortOrientation _
                           As XlSortOrientation = xlSortColumns)
' ==========================================================================
' Description : Sort an array.
'
' Parameters  : Arr             The array to sort
'               DimIdx          The Dimension to sort on
'                               in a multi-dimensional array,
'               SortOrder       Sort ascending or descending
'               SortOrientation Normally arrays are ordered by Columns
'                               ListBox lists are ordered by Rows
' ==========================================================================

    Const sPROC     As String = "BubbleSortArray"

    Dim bDone       As Boolean
    Dim bMultiDim   As Boolean

    Dim lCol        As Long
    Dim lDimCt      As Long
    Dim lIdx        As Long

    Dim vTmp        As Variant


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Determine if a multi-dim array
    ' ------------------------------
    lDimCt = NumberOfDimensions(Arr)
    bMultiDim = (lDimCt > 1)

    ' Determine the sort direction
    ' ----------------------------
    If (SortOrder = xlDescending) Then
        If (bMultiDim And (SortOrientation = xlSortRows)) Then
            GoTo SORT_DESC_MULTI_ROWS
        ElseIf bMultiDim Then
            GoTo SORT_DESC_MULTI
        Else
            GoTo SORT_DESCENDING
        End If
    Else
        If (bMultiDim And (SortOrientation = xlSortRows)) Then
            GoTo SORT_ASC_MULTI_ROWS
        ElseIf bMultiDim Then
            GoTo SORT_ASC_MULTI
        End If
    End If

    ' ----------------------------------------------------------------------
    ' Ascending
    ' ---------

    Do
        ' Assume we're finished
        ' ---------------------
        bDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For lIdx = LBound(Arr) To UBound(Arr) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(lIdx) > Arr(lIdx + 1)) Then

                ' ... swap them ...
                ' -----------------
                vTmp = Arr(lIdx)
                Arr(lIdx) = Arr(lIdx + 1)
                Arr(lIdx + 1) = vTmp

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                bDone = False
                vTmp = Empty
            End If

        Next lIdx

    Loop While Not bDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_ASC_MULTI:

    Do
        ' Assume we're finished
        ' ---------------------
        bDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For lIdx = LBound(Arr, 1) To UBound(Arr, 1) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(DimIdx, lIdx) > Arr(DimIdx, lIdx + 1)) Then

                ' ... swap them ...
                ' -----------------
                For lCol = 0 To lDimCt - 1
                    vTmp = Arr(lCol, lIdx)
                    Arr(lCol, lIdx) = Arr(lCol, lIdx + 1)
                    Arr(lCol, lIdx + 1) = vTmp
                    vTmp = Empty
                Next lCol

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                bDone = False
                vTmp = Empty
            End If

        Next lIdx

    Loop While Not bDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_ASC_MULTI_ROWS:

    Do
        ' Assume we're finished
        ' ---------------------
        bDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For lIdx = LBound(Arr, 1) To UBound(Arr, 1) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(lIdx, DimIdx) > Arr(lIdx + 1, DimIdx)) Then

                ' ... swap them ...
                ' -----------------
                For lCol = LBound(Arr, lDimCt) To UBound(Arr, lDimCt)
                    vTmp = Arr(lIdx, lCol)
                    Arr(lIdx, lCol) = Arr(lIdx + 1, lCol)
                    Arr(lIdx + 1, lCol) = vTmp
                    vTmp = Empty
                Next lCol

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                bDone = False
                vTmp = Empty
            End If

        Next lIdx

    Loop While Not bDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_DESCENDING:

    Do
        ' Assume we're finished
        ' ---------------------
        bDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For lIdx = LBound(Arr) To UBound(Arr) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(lIdx) < Arr(lIdx + 1)) Then

                ' ... swap them ...
                ' -----------------
                vTmp = Arr(lIdx)
                Arr(lIdx) = Arr(lIdx + 1)
                Arr(lIdx + 1) = vTmp

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                bDone = False
                vTmp = Empty
            End If

        Next lIdx

    Loop While Not bDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_DESC_MULTI:

    Do
        ' Assume we're finished
        ' ---------------------
        bDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For lIdx = LBound(Arr, 1) To UBound(Arr, 1) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(DimIdx, lIdx) < Arr(DimIdx, lIdx + 1)) Then

                ' ... swap them ...
                ' -----------------
                For lCol = 0 To lDimCt - 1
                    vTmp = Arr(lCol, lIdx)
                    Arr(lCol, lIdx) = Arr(lCol, lIdx + 1)
                    Arr(lCol, lIdx + 1) = vTmp
                    vTmp = Empty
                Next lCol

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                bDone = False
                vTmp = Empty
            End If

        Next lIdx

    Loop While Not bDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_DESC_MULTI_ROWS:

    Do
        ' Assume we're finished
        ' ---------------------
        bDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For lIdx = LBound(Arr, 1) To UBound(Arr, 1) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(lIdx, DimIdx) < Arr(lIdx + 1, DimIdx)) Then

                ' ... swap them ...
                ' -----------------
                For lCol = LBound(Arr, lDimCt) To UBound(Arr, lDimCt)
                    vTmp = Arr(lIdx, lCol)
                    Arr(lIdx, lCol) = Arr(lIdx + 1, lCol)
                    Arr(lIdx + 1, lCol) = vTmp
                    vTmp = Empty
                Next lCol

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                bDone = False
                vTmp = Empty
            End If

        Next lIdx

    Loop While Not bDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub

Public Function CollectionToArray(ByRef Col As VBA.Collection, _
                         Optional ByVal Base As Long) As Variant
' ==========================================================================
' Description : Convert a collection to an array
'
' Parameters  : Col     The collection to convert
'
' Returns     : Variant
' ==========================================================================

    Const sPROC As String = "CollectionToArray"

    Dim lIdx    As Long
    Dim vRtn    As Variant


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    vRtn = Array()
    ReDim vRtn(Base To Base + Col.Count - 1)

    For lIdx = 1 To Col.Count
        vRtn(lIdx + Base - 1) = Col.Item(lIdx)
    Next lIdx

    ' ----------------------------------------------------------------------

PROC_EXIT:

    CollectionToArray = vRtn

    On Error Resume Next
    Erase vRtn
    vRtn = Empty

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function CombineArrays(ParamArray Params() As Variant) As Variant
' ==========================================================================
' Description : Combine the contents of multiple arrays into a new array
'
' Parameters  : ParamArray  The arrays to combine
'
' Returns     : Variant (containing an array)
' ==========================================================================

    Const sPROC As String = "CombineArrays"

    Dim lIdx    As Long: lIdx = -1

    Dim vElt    As Variant
    Dim vAry    As Variant
    Dim vRtn    As Variant


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    vRtn = Array()

    For Each vAry In Params
        If IsArray(vAry) Then
            For Each vElt In vAry
                lIdx = lIdx + 1
                ReDim Preserve vRtn(0 To lIdx)
                vRtn(lIdx) = vElt
            Next vElt
        Else
            lIdx = lIdx + 1
            ReDim Preserve vRtn(0 To lIdx)
            vRtn(lIdx) = vAry
        End If
    Next vAry

    ' ----------------------------------------------------------------------

PROC_EXIT:

    CombineArrays = vRtn

    On Error Resume Next
    Erase vAry
    Erase vRtn
    vAry = Empty
    vRtn = Empty

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function GetArrayValue(ByRef Arr As Variant, _
                              ByVal XValue As String, _
                              ByVal YValue As String) As Variant
' ==========================================================================
' Description : Find data in a multi-dimensional array, and return it based
'                on the the X-Value (the header) and the Y-Value (the row).
'
' Parameters  : Arr       The data array to search
'               XValue    The header value. The first index in each dim.
'               YValue    The row index value to look for
'
' Returns     : Variant
' ==========================================================================

    Const sPROC As String = "GetArrayValue"

    Dim vRtn    As Variant

    Dim lIdxX   As Long
    Dim lIdxY   As Long
    Dim lLbx    As Long
    Dim lUBX    As Long
    Dim lLBY    As Long
    Dim lUBY    As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, YValue & " (" & XValue & ")")

    ' ----------------------------------------------------------------------
    ' Quit if no X or Y value provided
    ' --------------------------------

    If (Len(Trim(XValue)) = 0) Then
        GoTo PROC_EXIT
    End If

    If (Len(Trim(YValue)) = 0) Then
        GoTo PROC_EXIT
    End If

    ' ----------------------------------------------------------------------
    ' Find the column
    ' ---------------
    lLbx = LBound(Arr, 1)
    lUBX = UBound(Arr, 1)

    lLBY = LBound(Arr, 2)
    lUBY = UBound(Arr, 2)

    ' Find the header
    ' ---------------
    For lIdxX = lLbx To lUBX

        If (Arr(lIdxX, lLBY) = XValue) Then

            ' Find the row
            ' ------------
            For lIdxY = lLBY To lUBY
                If (Arr(lLbx, lIdxY) = YValue) Then
                    vRtn = Arr(lIdxX, lIdxY)
                    lIdxY = lUBY
                End If
            Next lIdxY

            lIdxX = lUBX
        End If

    Next lIdxX

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetArrayValue = vRtn

    Call Trace(tlMaximum, msMODULE, sPROC, vRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function GetElementIndex(ByRef Arr As Variant, _
                                ByVal Element As Variant, _
                       Optional ByVal Dimensions As Long = 1, _
                       Optional ByVal Dimension As Long) As Long
' ==========================================================================
' Description : Locate an element in an array
'
' Parameters  : Arr         The array to search
'               Element     The element to search for
'               Dimensions  The number of dimensions in the array
'               Dimension   The dimension to search in
'
' Returns     : Long
' ==========================================================================

    Const sPROC As String = "GetElementIndex"

    Dim lLB     As Long
    Dim lUB     As Long
    Dim lIdx    As Long
    Dim lRtn    As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Element)

    ' ----------------------------------------------------------------------

    If (Dimensions > 1) Then
        lLB = LBound(Arr, Dimensions)
        lUB = UBound(Arr, Dimensions)

        For lIdx = lLB To lUB
            If (Arr(Dimension, lIdx) = Element) Then
                lRtn = lIdx
                Exit For
            End If
        Next lIdx

    Else
        lLB = LBound(Arr)
        lUB = UBound(Arr)

        For lIdx = lLB To lUB
            If (Arr(lIdx) = Element) Then
                lRtn = lIdx
                Exit For
            End If
        Next lIdx
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetElementIndex = lRtn

    Call Trace(tlMaximum, msMODULE, sPROC, lRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function GetElementIndexX(ByRef Arr As Variant, _
                                 ByVal Element As Variant, _
                                 ByVal YIndex As Long) As Long
' ==========================================================================
' Description : Locate an element in an array across the X-axis
'
' Parameters  : Arr       The array to search
'               Element   The element to search for
'               YIndex    The Y-index to search in
'
' Returns     : Long
' ==========================================================================

    Const sPROC As String = "GetElementIndexX"

    Dim lLB     As Long
    Dim lUB     As Long
    Dim lIdx    As Long
    Dim lRtn    As Long


    On Error GoTo PROC_ERR
    Call Trace(tlVerbose, msMODULE, sPROC, Element)

    ' ----------------------------------------------------------------------

    lLB = LBound(Arr, 1)
    lUB = UBound(Arr, 1)

    For lIdx = lLB To lUB
        If (Arr(lIdx, YIndex) = Element) Then
            lRtn = lIdx
            Exit For
        End If
    Next lIdx

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetElementIndexX = lRtn

    Call Trace(tlVerbose, msMODULE, sPROC, lRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function InsertElement(ByRef Arr As Variant, _
                              ByVal Element As Variant, _
                     Optional ByVal Index As Long = -1) As Boolean
' ==========================================================================
' Description : Insert a value into an array.
'               The array must be dynamic, and everything after the
'               Index is moved down the array by 1.
'
' Parameters  : Arr       The array to modify
'               Element   The value to insert
'               Index     If Index is less than LBound or
'                         greater than UBound, the Element
'                         will be added to the end
'
' Returns     : Boolean   Returns True if successful
' ==========================================================================

    Const sPROC As String = "InsertElement"

    Dim bRtn    As Boolean

    Dim lIdx    As Long
    Dim lLB     As Long
    Dim lUB     As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Make sure it is an array
    ' ------------------------
    If (Not IsArray(Arr)) Then
        GoTo PROC_EXIT
    End If

    ' Only single-dimension
    ' arrays are allowed
    ' ---------------------
    If (NumberOfDimensions(Arr) <> 1) Then
        GoTo PROC_EXIT
    End If

    ' Assume success
    ' --------------
    bRtn = True

    ' Get the extents of the array
    ' ----------------------------
    lLB = LBound(Arr)
    lUB = UBound(Arr)

    ' If Index is less than LBound or greater than
    ' UBound, the Element will be added to the end
    ' --------------------------------------------
    If ((Index < lLB) Or (Index > lUB)) Then
        Index = lUB + 1
        ReDim Preserve Arr(lLB To Index)
        Arr(Index) = Element
        GoTo PROC_EXIT
    End If

    ' The Index is within the array.
    ' Move all elements after the index.
    ' ----------------------------------
    For lIdx = lUB To Index + 1 Step -1
        Arr(lIdx) = Arr(lIdx - 1)
    Next lIdx

    Arr(Index) = Element

    ' ----------------------------------------------------------------------

PROC_EXIT:

    InsertElement = bRtn


    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    bRtn = False

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function IsAllocated(ByRef Arr As Variant) As Boolean
' ==========================================================================
' Purpose   : Determines if an array is allocated
'
' Arguments : Arr     The array to test
'
' Returns   : Boolean Returns True if the array is allocated (either a
'                     static array or a dynamic array that has been sized
'                     with ReDim) or False if the array has not been
'                     allocated (a dynamic that has not yet been sized with
'                     ReDim, or a dynamic array that has been Erased).
' ==========================================================================

    Const sPROC As String = "IsAllocated"

    Dim bRtn    As Boolean
    Dim lUB     As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Quit if Arr is not an array
    ' ---------------------------
    If Not IsArray(Arr) Then
        bRtn = False
        GoTo PROC_EXIT
    End If

    ' Test the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occurred.
    ' ------------------------------------------------------------------
    On Error Resume Next                ' Err.Clear automatically invoked
    lUB = UBound(Arr, 1)

    If (Err.Number = ERR_SUCCESS) Then  ' No error - array is allocated
        bRtn = True

    Else                                ' Array is unallocated
        bRtn = False
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsAllocated = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, bRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function IsEmptyArray(ByRef Arr As Variant) As Boolean
' ==========================================================================
' Description : Determines if the array is empty (unallocated)
'               The VBA IsArray function indicates whether a variable is an
'               array, but it does not distinguish between allocated and
'               unallocated arrays. It will return TRUE for both
'               allocated and unallocated arrays. This function tests
'               whether the array has actually been allocated.
'
' Parameters  : Arr   The array to test
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "IsEmptyArray"

    Dim bRtn    As Boolean

    Dim lLB     As Long
    Dim lUB     As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Err.Clear
    On Error Resume Next

    ' Not an array
    ' ------------
    If (Not IsArray(Arr)) Then
        bRtn = True
        GoTo PROC_EXIT
    End If

    ' If not allocated an error will occur
    ' ------------------------------------
    lUB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        bRtn = True
        GoTo PROC_EXIT
    End If

    ' A new array has
    ' LBound = 0 and UBound = -1
    ' --------------------------
    Err.Clear
    lLB = LBound(Arr)
    If (lLB > lUB) Then
        bRtn = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsEmptyArray = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function IsInArray(ByRef Arr As Variant, _
                          ByVal Val As Variant) As Boolean
' ==========================================================================
' Description : Determines if a value is in an array
'
' Parameters  : Arr     The array to test
'               Val     The value to look for
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "IsInArray"

    Dim bRtn    As Boolean

    Dim lIdx    As Long
    Dim lLB     As Long
    Dim lUB     As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    lLB = LBound(Arr)
    lUB = UBound(Arr)

    For lIdx = lLB To lUB
        If (Arr(lIdx) = Val) Then
            bRtn = True
            Exit For
        End If
    Next

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsInArray = bRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Sub ListArray(ParamArray Params() As Variant)
' ==========================================================================
' Description : List the contents of an array in the Immediate Window
'
' Parameters  : ParamArray    The array to list
' ==========================================================================

    Const sPROC     As String = "ListArray"

    Dim lIdx        As Long: lIdx = -1

    Dim vParam      As Variant
    Dim vElement    As Variant


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Debug.Print String$(glLIST_LINELEN, gsLIST_LINECHAR)

    For Each vParam In Params

        ' Parse the array
        ' ---------------
        If IsArray(vParam) Then
            If (lIdx = -1) Then
                lIdx = LBound(vParam) - 1
            End If
            For Each vElement In vParam
                lIdx = lIdx + 1
                Debug.Print lIdx & " = ", vElement
            Next vElement

            ' List singleton item
            ' -------------------
        Else
            lIdx = lIdx + 1
            Debug.Print lIdx & " = ", vElement
        End If
    Next vParam

    Debug.Print String$(glLIST_LINELEN, gsLIST_LINECHAR)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, lIdx)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub

Public Function NumberOfDimensions(ByRef Arr As Variant) As Long
' ==========================================================================
' Purpose   : Determine the number of dimensions to an array
'
' Arguments : Arr     The array to test
'
' Returns   : Long    The number of dimensions in the array
' ==========================================================================

    Dim lIdx    As Long
    Dim lRtn    As Long
    Dim lUB     As Long


    On Error Resume Next

    ' ----------------------------------------------------------------------
    ' Increase the array index until an error occurs.
    ' An error will occur when it exceeds
    ' the number of dimensions in the array.
    ' -----------------------------------------------
    Err.Clear

    Do
        lIdx = lIdx + 1
        lUB = UBound(Arr, lIdx)
    Loop Until Err.Number <> 0

    lRtn = lIdx - 1

    ' ----------------------------------------------------------------------

PROC_EXIT:

    NumberOfDimensions = lRtn

End Function

Public Function NumberOfElements(ByRef Arr As Variant, _
                        Optional ByVal Dimension As Long = 1) As Long
' ==========================================================================
' Description : Determine the number of elements in an array dimension
'
' Parameters  : Arr         The array to examine
'               Dimension   The Dimension in the array. If Dimension
'                           is not provided, the first dimension is used.
'                           This function will return 0
'                           under the following circumstances:
'                             Arr is not an array
'                             Array is unallocated
'                             Dimension is less than 1
'                             Dimension is greater than the number of Dims
'
' Returns     : Long
' ==========================================================================

    Const sPROC As String = "NumberOfElements"

    Dim lRtn    As Long
    Dim lDimCt  As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Not an array
    ' ------------
    If (Not IsArray(Arr)) Then
        GoTo PROC_EXIT
    End If

    ' Array is unallocated
    ' --------------------
    If IsEmptyArray(Arr) Then
        GoTo PROC_EXIT
    End If

    ' Dimension is less than 1
    ' ------------------------
    If Dimension < 1 Then
        GoTo PROC_EXIT
    End If

    ' Get the number of dimensions
    ' ----------------------------
    lDimCt = NumberOfDimensions(Arr)

    ' Dimension greater than number of Dimensions
    ' -------------------------------------------
    If (lDimCt < Dimension) Then
        GoTo PROC_EXIT
    End If

    ' Get the number of elements
    ' --------------------------
    lRtn = UBound(Arr, Dimension) - LBound(Arr, Dimension) + 1

    ' ----------------------------------------------------------------------

PROC_EXIT:

    NumberOfElements = lRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function OneDimToTwo(ByRef Arr As Variant, _
                   Optional ByVal DataDim As Long = 1, _
                   Optional ByVal BaseDim1 As Long, _
                   Optional ByVal BaseDim2 As Long) As Variant
' ==========================================================================
' Description : Convert a 1-dimensional array to 2-dimensional
'
' Parameters  : Arr   The array to convert
'               DataDim   The dimension to copy the data to
'               BaseDim1  Specifies the lower bound of the first dimension
'               BaseDim2  Specifies the lower bound of the second dimension
'
' Notes       : Bases are usually 0 or 1 (default is zero), but can be
'               higher to align the index with a column or row
'
' Returns     : Variant
' ==========================================================================

    Const sPROC As String = "OneDimToTwo"

    Dim lIdx    As Long
    Dim lLB     As Long
    Dim lUB     As Long

    Dim vRtn    As Variant


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the size of the source
    ' --------------------------
    lLB = LBound(Arr)
    lUB = UBound(Arr)

    ' Size the array
    ' --------------
    vRtn = Array()
    ReDim vRtn(BaseDim1 To BaseDim1 + 1, BaseDim2 To lUB + BaseDim2)

    ' Copy the array
    ' --------------
    If (DataDim = 1) Then
        For lIdx = lLB To lUB
            vRtn(BaseDim1 + DataDim, BaseDim2 + lIdx) = Arr(lIdx)
        Next lIdx
    Else
        For lIdx = lLB To lUB
            vRtn(BaseDim1, BaseDim2 + lIdx) = Arr(lIdx)
        Next lIdx
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    OneDimToTwo = vRtn

    On Error Resume Next
    Erase vRtn
    vRtn = Empty

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Sub QuickSortArray(ByRef Arr As Variant, _
                 Optional ByRef LB As Long = -2, _
                 Optional ByRef UB As Long = -2)
' ==========================================================================
' Description : Sort an array using the recursive QuickSort algorithm.
'               This version is a 'balanced' QuickSort, selecting the
'               middle-indexed item as the pivot point for comparison.
'               Using the middle point could also allow this to be
'               considered a Bucket sort.
'
' Parameters  : Arr     The array to sort.
'                       For performance, this argument is not checked to
'                       ensure it is actually an array before starting
'                       (because it is called recursively). Always ensure
'                       this is an array before starting using IsArray(Arr)
'                       to avoid potential problems.
'
'               LB      The LowerBound of the array to sort.
'               UB      The UpperBound of the array to sort.
'
'                       If these values are not passed, the extents are
'                       automatically checked. By default it is set
'                       to -2, as -1 is the LBound of an uninitialized
'                       array, and 0 is the LBound of an initialized array.
'                       These values are then passed to recursive calls.
' ==========================================================================

    Const sPROC As String = "QuickSortArray"

    Dim lIdxL   As Long
    Dim lIdxU   As Long
    Dim lIdxMid As Long

    Dim vPivot  As Variant


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the extents of the array
    ' ----------------------------

    If (LB = -2) Then
        LB = LBound(Arr)
    End If
    If (UB = -2) Then
        UB = UBound(Arr)
    End If

    ' Only sort if there are elements to sort on
    ' ------------------------------------------
    If (Not (LB < UB)) Then
        GoTo PROC_EXIT
    Else
        lIdxL = LB
        lIdxU = UB
    End If

    ' Use the middle of the array
    ' as the pivot value to compare
    ' -----------------------------
    lIdxMid = (LB + UB) \ 2
    vPivot = Arr(lIdxMid)

    Do
        Do While (Arr(lIdxL) < vPivot)
            lIdxL = lIdxL + 1
        Loop

        Do While Arr(lIdxU) > vPivot
            lIdxU = lIdxU - 1
        Loop

        If (lIdxL <= lIdxU) Then
            Call QuickSortSwap(Arr, lIdxL, lIdxU)
            lIdxL = lIdxL + 1
            lIdxU = lIdxU - 1
        End If
    Loop Until (lIdxL > lIdxU)

    ' Work on the smaller partition first
    ' -----------------------------------
    If (lIdxU <= lIdxMid) Then
        Call QuickSortArray(Arr, LB, lIdxU)
        Call QuickSortArray(Arr, lIdxL, UB)
    Else
        Call QuickSortArray(Arr, lIdxL, UB)
        Call QuickSortArray(Arr, LB, lIdxU)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub

Private Sub QuickSortSwap(ByRef Arr As Variant, _
                          ByRef IdxL As Long, _
                          ByRef IdxU As Long)
' ==========================================================================
' Description : Perform the swap operation for the QuickSort routine
'
' Parameters  : Arr     The array to swap items in.
'               IdxL    The index of the lower item to swap.
'               IdxU    The index of the upper item to swap.
' ==========================================================================

    Const sPROC As String = "QuickSortSwap"

    Dim vTemp   As Variant

    vTemp = Arr(IdxU)
    Arr(IdxU) = Arr(IdxL)
    Arr(IdxL) = vTemp

End Sub

Public Sub RemoveElement(ByRef Arr As Variant, _
                         ByVal Index As Long)
' ==========================================================================
' Description : Removes an element from an array of items
'
' Parameters  : Arr     The array to modify
'               Index   The location of the item to remove
' ==========================================================================

    Const sPROC As String = "RemoveElement"

    Dim lIdx    As Long
    Dim lLB     As Long
    Dim lUB     As Long


    '  On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, Index)

    ' ----------------------------------------------------------------------

    If Not IsArray(Arr) Then
        Err.Raise 13, , "Type Mismatch"
        Exit Sub
    End If

    ' ----------------------------------

    lLB = LBound(Arr)
    lUB = UBound(Arr)

    If ((Index < lLB) Or (Index > lUB)) Then
        Err.Raise 9, , "Subscript out of Range"
        Exit Sub
    End If

    For lIdx = Index To lUB - 1
        Arr(lIdx) = Arr(lIdx + 1)
    Next

    On Error GoTo PROC_ERR
    ReDim Preserve Arr(lLB To lUB - 1)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub

Public Function StringToArray(ByVal Str As String, _
                     Optional ByVal Delimiter As String = " ", _
                     Optional ByVal Limit As Long = -1, _
                     Optional ByVal Compare _
                                 As VbCompareMethod = vbTextCompare) _
             As Variant
' ==========================================================================
' Description : Convert a string to an array
'
' Parameters  : Str         The string to convert
'               Delimiter   The character that separates the parts.
'                           If this is a zero-length string (vbNullString),
'                           each character will be separated.
'               Limit       Limit the number of sub-strings to return.
'
' Returns     : Variant
' ==========================================================================

    Const sPROC As String = "StringToArray"

    Dim lIdx    As Long
    Dim lLen    As Long
    Dim lPos    As Long

    Dim vRtn    As Variant


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Standard method
    ' ---------------

    If (Len(Delimiter) > 0) Then
        vRtn = Split(Str, Delimiter, Limit, Compare)
        GoTo PROC_EXIT
    End If

    ' Split characters
    ' ----------------

    If (Limit > 0) Then
        lLen = Limit
    Else
        lLen = Len(Str)
    End If

    vRtn = Array()
    ReDim vRtn(0 To lLen - 1)

    For lPos = 1 To lLen
        lIdx = lPos - 1
        vRtn(lIdx) = Mid$(Str, lPos, 1)
    Next lPos

    ' ----------------------------------------------------------------------

PROC_EXIT:

    StringToArray = vRtn

    On Error Resume Next
    Erase vRtn
    vRtn = Empty

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function UniqueItemsInArray(ByRef Source As Variant, _
                          Optional ByVal Count As Boolean) As Variant
' ==========================================================================
' Purpose   : Returns the unique items within an array
'
' Arguments : Source  The array to return items from
'
'           : Count   If True, return the count of items.
'                     If false (default) or is missing,
'                     return an array of unique items.
'
' Returns   : Variant
' ==========================================================================

    Const sPROC As String = "UniqueItemsInArray"

    Dim bMatched As Boolean

    Dim lItemCt As Long
    Dim lIdx    As Long
    Dim lUB     As Long

    Dim vItems  As Variant  ' Array of items
    Dim Element As Variant


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Retain the base of the source
    ' -----------------------------

    vItems = Array()
    ReDim vItems(LBound(Source) To LBound(Source))
    lUB = LBound(Source) - 1

    ' Loop through the source data array
    ' ----------------------------------
    For Each Element In Source
        ' Reset the flag
        ' --------------
        bMatched = False

        ' Has the item been added?
        ' ------------------------
        For lIdx = LBound(vItems) To UBound(vItems)
            If (Element = vItems(lIdx)) Then
                bMatched = True
                Exit For
            End If
        Next lIdx

        ' If not in list, add the item
        ' ----------------------------
        If ((Not bMatched) And (Not IsEmpty(Element))) Then
            lItemCt = lItemCt + 1
            lUB = lUB + 1
            ReDim Preserve vItems(LBound(vItems) To lUB)
            vItems(lUB) = Element
        End If

    Next Element

    ' ----------------------------------------------------------------------

PROC_EXIT:

    If Count Then
        UniqueItemsInArray = CVar(lItemCt)
    Else
        UniqueItemsInArray = vItems
    End If

    On Error Resume Next

    ' Release the allocated memory
    ' ----------------------------
    Erase vItems
    vItems = Empty

    Call Trace(tlMaximum, msMODULE, sPROC, lItemCt)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function
