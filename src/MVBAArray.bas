Attribute VB_Name = "MVBAArray"
'@IgnoreModule ProcedureNotUsed, LineLabelNotUsed, ConstantNotUsed
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

Public Const ERR_SUCCESS                As Long = 0   ' Generic success code

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

    Const PROCEDURE_NAME As String = "ArrayToString"

    Dim returnValue    As String


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    returnValue = Join(Arr, Delimiter)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ArrayToString = returnValue

    'Call Trace(tlMaximum, msMODULE, sPROC, sRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Function

Public Sub BubbleSortArray(ByRef Arr As Variant, _
                           Optional ByVal DimIdx As Long, _
                           Optional ByVal SortOrder As XlSortOrder = xlAscending, _
                           Optional ByVal SortOrientation As XlSortOrientation = xlSortColumns)
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

    Const PROCEDURE_NAME        As String = "BubbleSortArray"

    Dim isDone                  As Boolean
    Dim isMultiDimensionalArray As Boolean

    Dim columnNumber            As Long
    Dim dimensionsCount         As Long
    Dim index                   As Long

    Dim tempArray               As Variant


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Determine if a multi-dim array
    ' ------------------------------
    dimensionsCount = NumberOfDimensions(Arr)
    isMultiDimensionalArray = (dimensionsCount > 1)

    ' Determine the sort direction
    ' ----------------------------
    If (SortOrder = xlDescending) Then
        If (isMultiDimensionalArray And (SortOrientation = xlSortRows)) Then
            GoTo SORT_DESC_MULTI_ROWS
        ElseIf isMultiDimensionalArray Then
            GoTo SORT_DESC_MULTI
        Else
            GoTo SORT_DESCENDING
        End If
    Else
        If (isMultiDimensionalArray And (SortOrientation = xlSortRows)) Then
            GoTo SORT_ASC_MULTI_ROWS
        ElseIf isMultiDimensionalArray Then
            GoTo SORT_ASC_MULTI
        End If
    End If

    ' ----------------------------------------------------------------------
    ' Ascending
    ' ---------

    Do
        ' Assume we're finished
        ' ---------------------
        isDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For index = LBound(Arr) To UBound(Arr) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(index) > Arr(index + 1)) Then

                ' ... swap them ...
                ' -----------------
                tempArray = Arr(index)
                Arr(index) = Arr(index + 1)
                Arr(index + 1) = tempArray

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                isDone = False
                tempArray = Empty
            End If

        Next index

    Loop While Not isDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_ASC_MULTI:

    Do
        ' Assume we're finished
        ' ---------------------
        isDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For index = LBound(Arr, 1) To UBound(Arr, 1) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(DimIdx, index) > Arr(DimIdx, index + 1)) Then

                ' ... swap them ...
                ' -----------------
                For columnNumber = 0 To dimensionsCount - 1
                    tempArray = Arr(columnNumber, index)
                    Arr(columnNumber, index) = Arr(columnNumber, index + 1)
                    Arr(columnNumber, index + 1) = tempArray
                    tempArray = Empty
                Next columnNumber

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                isDone = False
                tempArray = Empty
            End If

        Next index

    Loop While Not isDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_ASC_MULTI_ROWS:

    Do
        ' Assume we're finished
        ' ---------------------
        isDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For index = LBound(Arr, 1) To UBound(Arr, 1) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(index, DimIdx) > Arr(index + 1, DimIdx)) Then

                ' ... swap them ...
                ' -----------------
                For columnNumber = LBound(Arr, dimensionsCount) To UBound(Arr, dimensionsCount)
                    tempArray = Arr(index, columnNumber)
                    Arr(index, columnNumber) = Arr(index + 1, columnNumber)
                    Arr(index + 1, columnNumber) = tempArray
                    tempArray = Empty
                Next columnNumber

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                isDone = False
                tempArray = Empty
            End If

        Next index

    Loop While Not isDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_DESCENDING:

    Do
        ' Assume we're finished
        ' ---------------------
        isDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For index = LBound(Arr) To UBound(Arr) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(index) < Arr(index + 1)) Then

                ' ... swap them ...
                ' -----------------
                tempArray = Arr(index)
                Arr(index) = Arr(index + 1)
                Arr(index + 1) = tempArray

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                isDone = False
                tempArray = Empty
            End If

        Next index

    Loop While Not isDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_DESC_MULTI:

    Do
        ' Assume we're finished
        ' ---------------------
        isDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For index = LBound(Arr, 1) To UBound(Arr, 1) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(DimIdx, index) < Arr(DimIdx, index + 1)) Then

                ' ... swap them ...
                ' -----------------
                For columnNumber = 0 To dimensionsCount - 1
                    tempArray = Arr(columnNumber, index)
                    Arr(columnNumber, index) = Arr(columnNumber, index + 1)
                    Arr(columnNumber, index + 1) = tempArray
                    tempArray = Empty
                Next columnNumber

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                isDone = False
                tempArray = Empty
            End If

        Next index

    Loop While Not isDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

SORT_DESC_MULTI_ROWS:

    Do
        ' Assume we're finished
        ' ---------------------
        isDone = True

        ' Loop through the array
        ' and compare the values
        ' ----------------------
        For index = LBound(Arr, 1) To UBound(Arr, 1) - 1

            ' Compare the values.
            ' If they are the wrong order, ...
            ' --------------------------------
            If (Arr(index, DimIdx) < Arr(index + 1, DimIdx)) Then

                ' ... swap them ...
                ' -----------------
                For columnNumber = LBound(Arr, dimensionsCount) To UBound(Arr, dimensionsCount)
                    tempArray = Arr(index, columnNumber)
                    Arr(index, columnNumber) = Arr(index + 1, columnNumber)
                    Arr(index + 1, columnNumber) = tempArray
                    tempArray = Empty
                Next columnNumber

                ' ... and clear the
                ' flag to loop again.
                ' -------------------
                isDone = False
                tempArray = Empty
            End If

        Next index

    Loop While Not isDone

    GoTo PROC_EXIT

    ' ----------------------------------------------------------------------

PROC_EXIT:

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Sub

Public Function CollectionToArray(ByVal Col As VBA.Collection, _
                         Optional ByVal Base As Long) As Variant
' ==========================================================================
' Description : Convert a collection to an array
'
' Parameters  : Col     The collection to convert
'
' Returns     : Variant
' ==========================================================================

    Const PROCEDURE_NAME As String = "CollectionToArray"

    Dim index       As Long
    Dim returnValue As Variant


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    returnValue = Array()
    ReDim returnValue(Base To Base + Col.Count - 1)

    For index = 1 To Col.Count
        returnValue(index + Base - 1) = Col.Item(index)
    Next index

    ' ----------------------------------------------------------------------

PROC_EXIT:

    CollectionToArray = returnValue

    On Error Resume Next
    Erase returnValue
    'vRtn = Empty

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Function

Public Function CombineArrays(ParamArray Params() As Variant) As Variant
' ==========================================================================
' Description : Combine the contents of multiple arrays into a new array
'
' Parameters  : ParamArray  The arrays to combine
'
' Returns     : Variant (containing an array)
' ==========================================================================

    Const PROCEDURE_NAME    As String = "CombineArrays"

    Dim index               As Long: index = -1

    Dim vElt                As Variant
    Dim vAry                As Variant
    Dim returnValue         As Variant


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    returnValue = Array()

    For Each vAry In Params
        If IsArray(vAry) Then
            For Each vElt In vAry
                index = index + 1
                ReDim Preserve returnValue(0 To index)
                returnValue(index) = vElt
            Next vElt
        Else
            index = index + 1
            ReDim Preserve returnValue(0 To index)
            returnValue(index) = vAry
        End If
    Next vAry

    ' ----------------------------------------------------------------------

PROC_EXIT:

    CombineArrays = returnValue

    On Error Resume Next
    Erase vAry
    Erase returnValue
    'vAry = Empty
    'vRtn = Empty

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

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

    Const PROCEDURE_NAME    As String = "GetArrayValue"

    Dim returnValue         As Variant

    Dim indexX              As Long
    Dim indexY              As Long
    Dim lowerBoundX         As Long
    Dim upperBoundX         As Long
    Dim lowerBoundY         As Long
    Dim upperBoundY         As Long


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, YValue & " (" & XValue & ")")

    ' ----------------------------------------------------------------------
    ' Quit if no X or Y value provided
    ' --------------------------------

    If (Len(Trim$(XValue)) = 0) Then
        GoTo PROC_EXIT
    End If

    If (Len(Trim$(YValue)) = 0) Then
        GoTo PROC_EXIT
    End If

    ' ----------------------------------------------------------------------
    ' Find the column
    ' ---------------
    lowerBoundX = LBound(Arr, 1)
    upperBoundX = UBound(Arr, 1)

    lowerBoundY = LBound(Arr, 2)
    upperBoundY = UBound(Arr, 2)

    ' Find the header
    ' ---------------
    For indexX = lowerBoundX To upperBoundX

        If (Arr(indexX, lowerBoundY) = XValue) Then

            ' Find the row
            ' ------------
            For indexY = lowerBoundY To upperBoundY
                If (Arr(lowerBoundX, indexY) = YValue) Then
                    returnValue = Arr(indexX, indexY)
                    indexY = upperBoundY
                End If
            Next indexY

            indexX = upperBoundX
        End If

    Next indexX

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetArrayValue = returnValue

    'Call Trace(tlMaximum, msMODULE, sPROC, vRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

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

    Const PROCEDURE_NAME    As String = "GetElementIndex"

    Dim lowerBound          As Long
    Dim upperBound          As Long
    Dim index               As Long
    Dim returnValue         As Long


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, Element)

    ' ----------------------------------------------------------------------

    If (Dimensions > 1) Then
        lowerBound = LBound(Arr, Dimensions)
        upperBound = UBound(Arr, Dimensions)

        For index = lowerBound To upperBound
            If (Arr(Dimension, index) = Element) Then
                returnValue = index
                Exit For
            End If
        Next index

    Else
        lowerBound = LBound(Arr)
        upperBound = UBound(Arr)

        For index = lowerBound To upperBound
            If (Arr(index) = Element) Then
                returnValue = index
                Exit For
            End If
        Next index
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetElementIndex = returnValue

    'Call Trace(tlMaximum, msMODULE, sPROC, lRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

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

    Const PROCEDURE_NAME    As String = "GetElementIndexX"

    Dim lowerBound          As Long
    Dim upperBound          As Long
    Dim index               As Long
    Dim returnValue         As Long


    On Error GoTo PROC_ERR
    'Call Trace(tlVerbose, msMODULE, sPROC, Element)

    ' ----------------------------------------------------------------------

    lowerBound = LBound(Arr, 1)
    upperBound = UBound(Arr, 1)

    For index = lowerBound To upperBound
        If (Arr(index, YIndex) = Element) Then
            returnValue = index
            Exit For
        End If
    Next index

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetElementIndexX = returnValue

    'Call Trace(tlVerbose, msMODULE, sPROC, lRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Function

Public Function InsertElement(ByRef Arr As Variant, _
                              ByVal Element As Variant, _
                     Optional ByVal index As Long = -1) As Boolean
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

    Const PROCEDURE_NAME    As String = "InsertElement"

    Dim localIndex          As Long:    localIndex = index
    Dim returnValue         As Boolean

    Dim counter             As Long
    Dim lowerBound          As Long
    Dim upperBound          As Long


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

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
    returnValue = True

    ' Get the extents of the array
    ' ----------------------------
    lowerBound = LBound(Arr)
    upperBound = UBound(Arr)

    ' If Index is less than LBound or greater than
    ' UBound, the Element will be added to the end
    ' --------------------------------------------
    If ((localIndex < lowerBound) Or (localIndex > upperBound)) Then
        localIndex = upperBound + 1
        ReDim Preserve Arr(lowerBound To localIndex)
        Arr(localIndex) = Element
        GoTo PROC_EXIT
    End If

    ' The Index is within the array.
    ' Move all elements after the index.
    ' ----------------------------------
    For counter = upperBound To localIndex + 1 Step -1
        Arr(counter) = Arr(counter - 1)
    Next counter

    Arr(localIndex) = Element

    ' ----------------------------------------------------------------------

PROC_EXIT:

    InsertElement = returnValue

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    'bRtn = False

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

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

    Const PROCEDURE_NAME As String = "IsAllocated"

    Dim returnValue    As Boolean
    'Dim lUB     As Long


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Quit if Arr is not an array
    ' ---------------------------
    If Not IsArray(Arr) Then
        returnValue = False
        GoTo PROC_EXIT
    End If

    ' Test the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occurred.
    ' ------------------------------------------------------------------
    On Error Resume Next                ' Err.Clear automatically invoked
    'lUB = UBound(Arr, 1)

    returnValue = (Err.Number = ERR_SUCCESS)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsAllocated = returnValue

    'Call Trace(tlMaximum, msMODULE, sPROC, bRtn)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

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

    Const PROCEDURE_NAME    As String = "IsEmptyArray"

    Dim returnValue         As Boolean

    Dim lowerBound          As Long
    Dim upperBound          As Long


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    Err.Clear
    On Error Resume Next

    ' Not an array
    ' ------------
    If (Not IsArray(Arr)) Then
        returnValue = True
        GoTo PROC_EXIT
    End If

    ' If not allocated an error will occur
    ' ------------------------------------
    upperBound = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        returnValue = True
        GoTo PROC_EXIT
    End If

    ' A new array has
    ' LBound = 0 and UBound = -1
    ' --------------------------
    Err.Clear
    lowerBound = LBound(Arr)
    If (lowerBound > upperBound) Then
        returnValue = True
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsEmptyArray = returnValue

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

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

    Const PROCEDURE_NAME    As String = "IsInArray"

    Dim returnValue         As Boolean

    Dim index               As Long
    Dim lowerBound          As Long
    Dim upperBound          As Long


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    lowerBound = LBound(Arr)
    upperBound = UBound(Arr)

    For index = lowerBound To upperBound
        If (Arr(index) = Val) Then
            returnValue = True
            Exit For
        End If
    Next

    ' ----------------------------------------------------------------------

PROC_EXIT:

    IsInArray = returnValue

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Function

Public Sub ListArray(ParamArray Params() As Variant)
' ==========================================================================
' Description : List the contents of an array in the Immediate Window
'
' Parameters  : ParamArray    The array to list
' ==========================================================================

    Const PROCEDURE_NAME    As String = "ListArray"

    Dim index               As Long: index = -1

    Dim vParam              As Variant
    Dim vElement            As Variant


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------

    'Debug.Print String$(glLIST_LINELEN, gsLIST_LINECHAR)

    For Each vParam In Params

        ' Parse the array
        ' ---------------
        If IsArray(vParam) Then
            If (index = -1) Then
                index = LBound(vParam) - 1
            End If
            For Each vElement In vParam
                index = index + 1
                Debug.Print index & " = ", vElement
            Next vElement

            ' List singleton item
            ' -------------------
        Else
            index = index + 1
            Debug.Print index & " = ", vElement
        End If
    Next vParam

    'Debug.Print String$(glLIST_LINELEN, gsLIST_LINECHAR)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    'Call Trace(tlMaximum, msMODULE, sPROC, lIdx)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Sub

Public Function NumberOfDimensions(ByVal Arr As Variant) As Long
' ==========================================================================
' Purpose   : Determine the number of dimensions to an array
'
' Arguments : Arr     The array to test
'
' Returns   : Long    The number of dimensions in the array
' ==========================================================================

    Dim index       As Long
    Dim returnValue As Long
    '@Ignore VariableNotUsed
    Dim upperBound  As Long


    On Error GoTo ErrorHandler

    ' ----------------------------------------------------------------------
    ' Increase the array index until an error occurs.
    ' An error will occur when it exceeds
    ' the number of dimensions in the array.
    ' -----------------------------------------------
    Err.Clear

    Do
        index = index + 1
        upperBound = UBound(Arr, index)
    Loop Until Err.Number <> 0

    returnValue = index - 1

    ' ----------------------------------------------------------------------

PROC_EXIT:

    NumberOfDimensions = returnValue
    
    Exit Function
ErrorHandler:

    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
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

    Const PROCEDURE_NAME    As String = "NumberOfElements"

    Dim returnValue         As Long
    Dim dimensionsCount     As Long


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

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
    dimensionsCount = NumberOfDimensions(Arr)

    ' Dimension greater than number of Dimensions
    ' -------------------------------------------
    If (dimensionsCount < Dimension) Then
        GoTo PROC_EXIT
    End If

    ' Get the number of elements
    ' --------------------------
    returnValue = UBound(Arr, Dimension) - LBound(Arr, Dimension) + 1

    ' ----------------------------------------------------------------------

PROC_EXIT:

    NumberOfElements = returnValue

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Function

Public Function OneDimToTwo(ByRef Arr As Variant, _
                   Optional ByVal DataDim As Long = 1, _
                   Optional ByVal firstDimensionLowerBound As Long, _
                   Optional ByVal secondDimensionLowerBound As Long) As Variant
' ==========================================================================
' Description : Convert a 1-dimensional array to 2-dimensional
'
' Parameters  : Arr                         The array to convert
'               DataDim                     The dimension to copy the data to
'               firstDimensionLowerBound    Specifies the lower bound of the first dimension
'               secondDimensionLowerBound   Specifies the lower bound of the second dimension
'
' Notes       : Bases are usually 0 or 1 (default is zero), but can be
'               higher to align the index with a column or row
'
' Returns     : Variant
' ==========================================================================

    Const PROCEDURE_NAME        As String = "OneDimToTwo"

    Dim index                   As Long
    Dim inputArrayLowerBound    As Long
    Dim inputArrayUpperBound    As Long

    Dim returnValue             As Variant


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the size of the source
    ' --------------------------
    inputArrayLowerBound = LBound(Arr)
    inputArrayUpperBound = UBound(Arr)

    ' Size the array
    ' --------------
    returnValue = Array()
    ReDim returnValue(firstDimensionLowerBound To firstDimensionLowerBound + 1, secondDimensionLowerBound To inputArrayUpperBound + secondDimensionLowerBound)

    ' Copy the array
    ' --------------
    If (DataDim = 1) Then
        For index = inputArrayLowerBound To inputArrayUpperBound
            returnValue(firstDimensionLowerBound + DataDim, secondDimensionLowerBound + index) = Arr(index)
        Next index
    Else
        For index = inputArrayLowerBound To inputArrayUpperBound
            returnValue(firstDimensionLowerBound, secondDimensionLowerBound + index) = Arr(index)
        Next index
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    OneDimToTwo = returnValue

    On Error Resume Next
    Erase returnValue
    'vRtn = Empty

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Function

Public Sub QuickSortArray(ByRef Arr As Variant, _
                 Optional ByRef lowerBound As Long = -2, _
                 Optional ByRef upperBound As Long = -2)
' ===================================================================================
' Description : Sort an array using the recursive QuickSort algorithm.
'               This version is a 'balanced' QuickSort, selecting the
'               middle-indexed item as the pivot point for comparison.
'               Using the middle point could also allow this to be
'               considered a Bucket sort.
'
' Parameters  : Arr             The array to sort.
'                               For performance, this argument is not checked to
'                               ensure it is actually an array before starting
'                               (because it is called recursively). Always ensure
'                               this is an array before starting using IsArray(Arr)
'                               to avoid potential problems.
'
'               lowerBound      The LowerBound of the array to sort.
'               upperBound      The UpperBound of the array to sort.
'
'                               If these values are not passed, the extents are
'                               automatically checked. By default it is set
'                               to -2, as -1 is the LBound of an uninitialized
'                               array, and 0 is the LBound of an initialized array.
'                               These values are then passed to recursive calls.
' ===================================================================================

    Const PROCEDURE_NAME    As String = "QuickSortArray"

    Dim indexLowerBound     As Long
    Dim indexUpperBound     As Long
    Dim indexMid            As Long

    Dim vPivot              As Variant


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the extents of the array
    ' ----------------------------

    If (lowerBound = -2) Then
        lowerBound = LBound(Arr)
    End If
    If (upperBound = -2) Then
        upperBound = UBound(Arr)
    End If

    ' Only sort if there are elements to sort on
    ' ------------------------------------------
    If lowerBound >= upperBound Then GoTo PROC_EXIT
    
    indexLowerBound = lowerBound
    indexUpperBound = upperBound

    ' Use the middle of the array
    ' as the pivot value to compare
    ' -----------------------------
    indexMid = (lowerBound + upperBound) \ 2
    vPivot = Arr(indexMid)

    Do
        Do While (Arr(indexLowerBound) < vPivot)
            indexLowerBound = indexLowerBound + 1
        Loop

        Do While Arr(indexUpperBound) > vPivot
            indexUpperBound = indexUpperBound - 1
        Loop

        If (indexLowerBound <= indexUpperBound) Then
            QuickSortSwap Arr, indexLowerBound, indexUpperBound
            indexLowerBound = indexLowerBound + 1
            indexUpperBound = indexUpperBound - 1
        End If
    Loop Until (indexLowerBound > indexUpperBound)

    ' Work on the smaller partition first
    ' -----------------------------------
    If (indexUpperBound <= indexMid) Then
        QuickSortArray Arr, lowerBound, indexUpperBound
        QuickSortArray Arr, indexLowerBound, upperBound
    Else
        QuickSortArray Arr, indexLowerBound, upperBound
        QuickSortArray Arr, lowerBound, indexUpperBound
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Sub

'@Ignore ParameterCanBeByVal
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

    Const PROCEDURE_NAME    As String = "QuickSortSwap"

    Dim vTemp               As Variant

    vTemp = Arr(IdxU)
    Arr(IdxU) = Arr(IdxL)
    Arr(IdxL) = vTemp

End Sub

Public Sub RemoveElement(ByRef Arr As Variant, _
                         ByVal index As Long)
' ==========================================================================
' Description : Removes an element from an array of items
'
' Parameters  : Arr     The array to modify
'               Index   The location of the item to remove
' ==========================================================================

    Const PROCEDURE_NAME    As String = "RemoveElement"

    Dim counter             As Long
    Dim lowerBound          As Long
    Dim upperBound          As Long


    '  On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, Index)

    ' ----------------------------------------------------------------------

    If Not IsArray(Arr) Then
        Err.Raise 13, , "Type Mismatch"
        Exit Sub
    End If

    ' ----------------------------------

    lowerBound = LBound(Arr)
    upperBound = UBound(Arr)

    If ((index < lowerBound) Or (index > upperBound)) Then
        Err.Raise 9, , "Subscript out of Range"
        Exit Sub
    End If

    For counter = index To upperBound - 1
        Arr(counter) = Arr(counter + 1)
    Next

    On Error GoTo PROC_ERR
    ReDim Preserve Arr(lowerBound To upperBound - 1)

    ' ----------------------------------------------------------------------

PROC_EXIT:

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Sub

Public Function StringToArray(ByVal stringToSplit As String, _
                     Optional ByVal Delimiter As String = " ", _
                     Optional ByVal Limit As Long = -1, _
                     Optional ByVal Compare As VbCompareMethod = vbTextCompare) _
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

    Const PROCEDURE_NAME    As String = "StringToArray"

    Dim index               As Long
    Dim lengthToSplit       As Long
    Dim position            As Long

    Dim returnValue         As Variant


    On Error GoTo PROC_ERR
    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Standard method
    ' ---------------

    If (Len(Delimiter) > 0) Then
        returnValue = Split(stringToSplit, Delimiter, Limit, Compare)
        GoTo PROC_EXIT
    End If

    ' Split characters
    ' ----------------

    If (Limit > 0) Then
        lengthToSplit = Limit
    Else
        lengthToSplit = Len(stringToSplit)
    End If

    returnValue = Array()
    ReDim returnValue(0 To lengthToSplit - 1)

    For position = 1 To lengthToSplit
        index = position - 1
        returnValue(index) = Mid$(stringToSplit, position, 1)
    Next position

    ' ----------------------------------------------------------------------

PROC_EXIT:

    StringToArray = returnValue

    On Error Resume Next
    Erase returnValue
    'vRtn = Empty

    'Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

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

    Const PROCEDURE_NAME    As String = "UniqueItemsInArray"

    Dim isMatched           As Boolean

    Dim itemCount           As Long
    Dim index               As Long
    Dim upperBound          As Long

    Dim items               As Variant  ' Array of items
    Dim Element             As Variant


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    ' Retain the base of the source
    ' -----------------------------

    items = Array()
    ReDim items(LBound(Source) To LBound(Source))
    upperBound = LBound(Source) - 1

    ' Loop through the source data array
    ' ----------------------------------
    For Each Element In Source
        ' Reset the flag
        ' --------------
        isMatched = False

        ' Has the item been added?
        ' ------------------------
        For index = LBound(items) To UBound(items)
            If (Element = items(index)) Then
                isMatched = True
                Exit For
            End If
        Next index

        ' If not in list, add the item
        ' ----------------------------
        If ((Not isMatched) And (Not IsEmpty(Element))) Then
            itemCount = itemCount + 1
            upperBound = upperBound + 1
            ReDim Preserve items(LBound(items) To upperBound)
            items(upperBound) = Element
        End If

    Next Element

    ' ----------------------------------------------------------------------

PROC_EXIT:

    If Count Then
        UniqueItemsInArray = CVar(itemCount)
    Else
        UniqueItemsInArray = items
    End If

    On Error Resume Next

    ' Release the allocated memory
    ' ----------------------------
    Erase items
    'vItems = Empty

    'Call Trace(tlMaximum, msMODULE, sPROC, lItemCt)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

'    If ErrorHandler(msMODULE, sPROC) Then
'        Stop
'        Resume
'    Else
'        Resume PROC_EXIT
'    End If

End Function


