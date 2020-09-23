Attribute VB_Name = "mStableQuick"
Option Explicit                      ' -©Rd 2006-

' Rd's stable non-recursive quick sort algorithms

' You are free to use any part or all of this code
' even for commercial purposes in any way you wish
' under the one strict condition that no copyright
' notice is moved or removed from where it is.

' For comments, suggestions or bug reports you can
' contact me at rd•edwards•bigpond•com.

' Declare some CopyMemory Alias's (thanks Bruce :)
Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)

' More efficient repeated use of numeric literals
Private Const NEG1 = -1&, n0 = 0&, n1 = 1&, n2 = 2&, n3 = 3&, n4 = 4&, n5 = 5&
Private Const n6 = 6&, n7 = 7&, n8 = 8&, n12 = 12&, n16 = 16&, n32 = 32&

Private Enum SAFEATURES
    FADF_AUTO = &H1              ' Array is allocated on the stack
    FADF_STATIC = &H2            ' Array is statically allocated
    FADF_EMBEDDED = &H4          ' Array is embedded in a structure
    FADF_FIXEDSIZE = &H10        ' Array may not be resized or reallocated
    FADF_BSTR = &H100            ' An array of BSTRs
    FADF_UNKNOWN = &H200         ' An array of IUnknown*
    FADF_DISPATCH = &H400        ' An array of IDispatch*
    FADF_VARIANT = &H800         ' An array of VARIANTs
    FADF_RESERVED = &HFFFFF0E8   ' Bits reserved for future use
    #If False Then
        Dim FADF_AUTO, FADF_STATIC, FADF_EMBEDDED, FADF_FIXEDSIZE, FADF_BSTR, FADF_UNKNOWN, FADF_DISPATCH, FADF_VARIANT, FADF_RESERVED
    #End If
End Enum
Private Const VT_BYREF = &H4000& ' Tests whether the InitedArray routine was passed a Variant that contains an array, rather than directly an array in the former case ptr already points to the SA structure. Thanks to Monte Hansen for this fix

Private Type SAFEARRAY
    cDims       As Integer       ' Count of dimensions in this array (only 1 supported)
    fFeatures   As Integer       ' Bitfield flags indicating attributes of a particular array
    cbElements  As Long          ' Byte size of each element of the array
    cLocks      As Long          ' Number of times the array has been locked without corresponding unlock
    pvData      As Long          ' Pointer to the start of the array data (use only if cLocks > 0)
End Type
Private Type SABOUNDS            ' This module supports single dimension arrays only
    cElements As Long            ' Count of elements in this dimension
    lLBound   As Long            ' The lower-bounding index of this dimension
    lUBound   As Long            ' The upper-bounding index of this dimension
End Type

Private lbs() As Long, ubs() As Long ' Non-recursive quicksort and insert/binary hybrid stacks

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Enum eCompare
    Lesser = -1&
    Equal = 0&
    Greater = 1&
    #If False Then
        Dim Lesser, Equal, Greater
    #End If
End Enum

Public Enum eDirection
    dDescending = -1&
    dDefault = 0&
    dAscending = 1&
    #If False Then
        Dim dDescending, dDefault, dAscending
    #End If
End Enum

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const dDefault_Direction = dAscending
Private mComp As eCompare
Private mCaseSens As VbCompareMethod
Private mDirection As eDirection

Property Get CaseSensitivity() As VbCompareMethod
    CaseSensitivity = mCaseSens
End Property

Property Let CaseSensitivity(ByVal eNewMethod As VbCompareMethod)
    mCaseSens = eNewMethod
End Property

Property Get SortDirection() As eDirection
    If mDirection = Equal Then mDirection = dDefault_Direction
    SortDirection = mDirection
End Property

Property Let SortDirection(ByVal eNewDirection As eDirection)
    If eNewDirection = dDefault Then
        If mDirection = Equal Then mDirection = dDefault_Direction
    Else
        mDirection = eNewDirection
    End If
End Property

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Property Get SortCompare() As eCompare
    SortCompare = mComp
End Property

Private Property Let SortCompare(ByVal eNewCompare As eCompare)
    mComp = eNewCompare
End Property

' + Stable SwapSort +++++++++++++++++++++++++++++++++++++

' This is a non-recursive quick-sort based algorithm that has
' been written from the ground up as a stable alternative to
' the blindingly fast quick-sort.

' It is not quite as fast as the outright fastest non-stable
' quick-sort, but is still very fast as it uses buffers and
' copymemory and is beaten by none of my other string sorting
' algorithms except my non-stable quick-sorts.

' A standard quick-sort only moves items that need swapping,
' while this stable algorithm manipulates all items on every
' iteration to keep them all in relative positions to one
' another. This algorithm I have dubbed the Avalanche©.

' I'm not sure if this is still a true quick-sort; it could
' be considered a quick-bubble hybrid on api steriods :)

Private Sub strSwapStable(sA() As String, ByVal lbA As Long, ByVal ubA As Long)
    ' This is my stable non-recursive swap sort
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim lpS As Long, Idx As Long, pvt As Long
    Dim item As String, lpStr As Long
    Dim lA_1() As Long, lpL_1 As Long
    Dim lA_2() As Long, lpL_2 As Long

    Idx = ubA - lbA ' cnt-1
    ReDim lA_1(n0 To Idx) As Long
    ReDim lA_2(n0 To Idx) As Long
    lpL_1 = VarPtr(lA_1(n0))
    lpL_2 = VarPtr(lA_2(n0))

    pvt = (Idx \ n8) + n16             ' Allow for worst case senario + some
    ReDim lbs(n1 To pvt) As Long       ' Stack to hold pending lower boundries
    ReDim ubs(n1 To pvt) As Long       ' Stack to hold pending upper boundries
    lpStr = VarPtr(item)               ' Cache pointer to the string variable
    lpS = VarPtr(sA(lbA)) - (lbA * n4) ' Cache pointer to the string array

    Do: ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA         ' Get pivot index position
        CopyMemByV lpStr, lpS + (pvt * n4), n4 ' Grab current value into item

        For Idx = lbA To pvt - n1
            If (StrComp(sA(Idx), item, mCaseSens) = mComp) Then ' (idx > item)
                CopyMemByV lpL_1 + (ptr1 * n4), lpS + (Idx * n4), n4  '3
                ptr1 = ptr1 + n1
            Else
                CopyMemByV lpL_2 + (ptr2 * n4), lpS + (Idx * n4), n4  '1
                ptr2 = ptr2 + n1
            End If
        Next
        For Idx = pvt + n1 To ubA
            If (StrComp(item, sA(Idx), mCaseSens) = mComp) Then ' (item > idx)
                CopyMemByV lpL_2 + (ptr2 * n4), lpS + (Idx * n4), n4  '2
                ptr2 = ptr2 + n1
            Else
                CopyMemByV lpL_1 + (ptr1 * n4), lpS + (Idx * n4), n4  '4
                ptr1 = ptr1 + n1
            End If
        Next '-Avalanche ©Rd-
        CopyMemByV lpS + (lbA * n4), lpL_2, ptr2 * n4
        CopyMemByV lpS + ((lbA + ptr2) * n4), lpStr, n4 ' Re-assign current
        CopyMemByV lpS + ((lbA + ptr2 + n1) * n4), lpL_1, ptr1 * n4

        If (ptr2 > n1) Then
            If (ptr1 > n1) Then cnt = cnt + n1: lbs(cnt) = lbA + ptr2 + n1: ubs(cnt) = ubA
            ubA = lbA + ptr2 - n1
        ElseIf (ptr1 > n1) Then
            lbA = lbA + ptr2 + n1
        Else
            If cnt = n0 Then Exit Do
            lbA = lbs(cnt): ubA = ubs(cnt): cnt = cnt - n1
        End If
    Loop: Erase lbs: Erase ubs: CopyMem ByVal lpStr, 0&, n4
End Sub

Sub strSwapStabletemp(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CaseSensitive As VbCompareMethod, Optional ByVal Direction As eDirection = dDefault) '-©Rd 2006-
    ' Delegates to my Stable Non-Recursive Swap-Sort
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    If Direction = dDefault Then Direction = SortDirection
    SortCompare = Direction
    CaseSensitivity = CaseSensitive
    'strSwapStable sA, lbA, ubA
    strSwapStable2 sA, lbA, ubA
End Sub

' + Stable SwapSort Indexed Version ++++++++++++++

' This is a stable non-recursive indexed swap sort.

' It is noticably faster than the non-indexed version!

' It has the benifit of indexing which allows the source
' array to remain unchanged. This also allows the index
' array to be passed on to other sort processes to be
' further manipulated.

Private Sub strSwapStableIndexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    ' This is my indexed stable non-recursive swap sort
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim lpS As Long, Idx As Long, pvt As Long
    Dim item As String, lpStr As Long
    Dim lA_1() As Long, lpL_1 As Long
    Dim lA_2() As Long, lpL_2 As Long
    Dim idxItem As Long, lpI As Long

    Idx = ubA - lbA ' cnt-1
    ReDim lA_1(n0 To Idx) As Long
    ReDim lA_2(n0 To Idx) As Long
    lpL_1 = VarPtr(lA_1(n0))
    lpL_2 = VarPtr(lA_2(n0))

    pvt = (Idx \ n8) + n16             ' Allow for worst case senario + some
    ReDim lbs(n1 To pvt) As Long       ' Stack to hold pending lower boundries
    ReDim ubs(n1 To pvt) As Long       ' Stack to hold pending upper boundries
    lpStr = VarPtr(item)               ' Cache pointer to the string variable
    lpS = VarPtr(sA(lbA)) - (lbA * n4) ' Cache pointer to the string array

    Do: ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA ' Get pivot index position
        idxItem = idxA(pvt)            ' Grab current value into item
        CopyMemByV lpStr, lpS + (idxItem * n4), n4
        lpI = VarPtr(idxA(lbA))        ' Cache pointer to the index array

        For Idx = lbA To pvt - n1
            If (StrComp(sA(idxA(Idx)), item, mCaseSens) = mComp) Then ' (idx > item)
                lA_1(ptr1) = idxA(Idx) '3
                ptr1 = ptr1 + n1
            Else
                lA_2(ptr2) = idxA(Idx) '1
                ptr2 = ptr2 + n1
            End If
        Next
        For Idx = pvt + n1 To ubA
            If (StrComp(item, sA(idxA(Idx)), mCaseSens) = mComp) Then ' (item > idx)
                lA_2(ptr2) = idxA(Idx) '2
                ptr2 = ptr2 + n1
            Else
                lA_1(ptr1) = idxA(Idx) '4
                ptr1 = ptr1 + n1
            End If
        Next '-Avalanche ©Rd-
        CopyMemByV lpI, lpL_2, ptr2 * n4
        idxA(lbA + ptr2) = idxItem
        CopyMemByV lpI + ((ptr2 + n1) * n4), lpL_1, ptr1 * n4

        If (ptr2 > n1) Then
            If (ptr1 > n1) Then cnt = cnt + n1: lbs(cnt) = lbA + ptr2 + n1: ubs(cnt) = ubA
            ubA = lbA + ptr2 - n1
        ElseIf (ptr1 > n1) Then
            lbA = lbA + ptr2 + n1
        Else
            If cnt = n0 Then Exit Do
            lbA = lbs(cnt): ubA = ubs(cnt): cnt = cnt - n1
        End If
    Loop: Erase lbs: Erase ubs: CopyMem ByVal lpStr, 0&, n4
End Sub

' These results are the fastest times produced by each
' routine for each operation, tested on my 866MHz P3.

' SwapStable
' It can sort 99,996 unsorted strings in 0.7435 seconds!
' It can re-sort 99,996 sorted strings in 0.5182 seconds!
' It can rev-sort 99,996 sorted strings in 0.5706 seconds!
' It can pretty-sort 99,996 sorted strings in 1.9488 seconds!
' It can rev-pretty 99,996 sorted strings in 2.0740 seconds!

' SwapStable Indexed
' It can sort 99,996 unsorted strings in 0.6936 seconds!
' It can re-sort 99,996 sorted strings in 0.4335 seconds!
' It can rev-sort 99,996 sorted strings in 0.4728 seconds!
' It can pretty-sort 99,996 sorted strings in 1.6907 seconds!
' It can rev-pretty 99,996 sorted strings in 1.7980 seconds!

' SwapStable2
' It can sort 99,996 unsorted strings in 0.7289 seconds!
' It can re-sort 99,996 sorted strings in 0.5202 seconds!
' It can rev-sort 99,996 sorted strings in 0.5569 seconds!
' It can pretty-sort 99,996 sorted strings in 1.5674 seconds!
' It can rev-pretty 99,996 sorted strings in 1.5944 seconds!

' SwapStable2 Indexed
' It can sort 99,996 unsorted strings in 0.6855 seconds!
' It can re-sort 99,996 sorted strings in 0.4350 seconds!
' It can rev-sort 99,996 sorted strings in 0.4649 seconds!
' It can pretty-sort 99,996 sorted strings in 1.3631 seconds!
' It can rev-pretty 99,996 sorted strings in 1.3947 seconds!

' These results are the fastest times produced by each routine
' when sorting 249,999 unsorted strings, tested on my 866MHz P3.

' Non-Stable SwapSort v4 . . 1.4704 seconds
' Non-Stable TriQuickSort. . 1.7481 seconds
' SwapStable v2.1 Indexed. . 1.9576 seconds
' SwapStable v2.1. . . . . . 1.9617 seconds
' Stable Twister v5. . . . . 2.1439 seconds
' Non-Stable HeapSort v1 . . 2.7836 seconds
' Non-Stable ShellHybrid . . 3.8835 seconds
' Non-Stable Shaffer QSort . 5.4572 seconds

Sub strSwapStableIdxtemp(sA() As String, lIdxA() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CaseSensitive As VbCompareMethod, Optional ByVal Direction As eDirection = dDefault) '-©Rd 2006-
    ' Delegates to my Indexed Stable Non-Recursive Swap-Sort
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    If Direction = dDefault Then Direction = SortDirection
    SortCompare = Direction
    CaseSensitivity = CaseSensitive
    ValidateIndexArray lIdxA, lbA, ubA
    'strSwapStableIndexed sA, lIdxA, lbA, ubA
    strSwapStable2Indexed sA, lIdxA, lbA, ubA
End Sub

' + Stable SwapSort v2 ++++++++++++++++++++++++++++++++

' This is a re-working of my stable swap sort algorithm.

' A runner section has been added to handle a very hard job
' for a stable sorter - reverse-sorting case-insensitively
' on data pre-sorted case-sensitively (lower-case first in
' dAscending order, or capitals first in dDescending order).

' It utilises a runner technique to boost this very demanding
' operation - down from 2.0 to 1.5 seconds on 100,000 items.

' Adding runners has also boosted stable reverse-sorting and
' same-direction pretty sorting operations. If performing a
' reverse or pretty sort operation the code can identify this
' state and the runners are turned on automatically.

' It identifies when the avalanche process is producing a zero
' count buffer one way and so is moving all items the other way,
' indicating that the data is in a pre-sorted state.

' The range becomes very small on unsorted data before a small
' range produces a zero count buffer (shifting no items up/down
' in relation to the current item) while it is recognisd almost
' immediately on reverse-sorted data, and reasonably quickly on
' reverse-pretty and same-direction pretty sorting operations.

' Note also that this stable reverse-soring operation is quite
' different to a non-stable inversion style reverse operation.

Private Sub strSwapStable2(sA() As String, ByVal lbA As Long, ByVal ubA As Long)
    ' This is an even faster stable non-recursive swap sort
    Dim item As String, lpStr As Long, lpS As Long
    Dim Idx As Long, optimal As Long, pvt As Long
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim base As Long, run As Long, cast As Long
    Dim inter1 As Long, inter2 As Long
    Dim lA_1() As Long, lpL_1 As Long
    Dim lA_2() As Long, lpL_2 As Long
    Dim lPrettyReverse As Long

    Idx = ubA - lbA ' cnt-1
    ReDim lA_1(n0 To Idx) As Long
    ReDim lA_2(n0 To Idx) As Long
    lpL_1 = VarPtr(lA_1(n0))
    lpL_2 = VarPtr(lA_2(n0))

    pvt = (Idx \ n8) + n16             ' Allow for worst case senario + some
    ReDim lbs(n1 To pvt) As Long       ' Stack to hold pending lower boundries
    ReDim ubs(n1 To pvt) As Long       ' Stack to hold pending upper boundries
    lpStr = VarPtr(item)               ' Cache pointer to the string variable
    lpS = VarPtr(sA(lbA)) - (lbA * n4) ' Cache pointer to the string array

    Do: ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA         ' Get pivot index position
        CopyMemByV lpStr, lpS + (pvt * n4), n4 ' Grab current value into item
        For Idx = lbA To pvt - n1
            If (StrComp(sA(Idx), item, mCaseSens) = mComp) Then ' (idx > item)
                CopyMemByV lpL_2 + (ptr2 * n4), lpS + (Idx * n4), n4  '3
                ptr2 = ptr2 + n1
            Else
                CopyMemByV lpL_1 + (ptr1 * n4), lpS + (Idx * n4), n4  '1
                ptr1 = ptr1 + n1
            End If
        Next
        inter1 = ptr1: inter2 = ptr2
        For Idx = pvt + n1 To ubA
            If (StrComp(item, sA(Idx), mCaseSens) = mComp) Then ' (idx < item)
                CopyMemByV lpL_1 + (ptr1 * n4), lpS + (Idx * n4), n4  '2
                ptr1 = ptr1 + n1
            Else
                CopyMemByV lpL_2 + (ptr2 * n4), lpS + (Idx * n4), n4  '4
                ptr2 = ptr2 + n1
            End If
        Next '-Avalanche v2 ©Rd-
        CopyMemByV lpS + (lbA * n4), lpL_1, ptr1 * n4
        CopyMemByV lpS + ((lbA + ptr1) * n4), lpStr, n4 ' Re-assign current
        CopyMemByV lpS + ((lbA + ptr1 + n1) * n4), lpL_2, ptr2 * n4

        If lPrettyReverse <> n0 Then
        ElseIf (ubA - lbA < 250&) Then ' Ignore false indicators
        ElseIf (inter1 = n0) Then
            If (inter2 = ptr2) Then    ' Reverse
                lPrettyReverse = 10000
            ElseIf (ptr1 = n0) Then    ' Pretty
                lPrettyReverse = 50000
            End If
        ElseIf (inter2 = n0) Then      ' Pretty
            If (ptr2 = n0) Then
                lPrettyReverse = 50000
        End If: End If

        If (lPrettyReverse <> n0) Then
            If (ptr1 - inter1 <> n0) And (inter1 < lPrettyReverse) Then  ' Runners dislike super large ranges
                CopyMemByV lpStr, lpS + ((lbA + ptr1 - n1) * n4), n4
                optimal = lbA + (inter1 \ n2)
                run = lbA + inter1
                Do While run > optimal                                   ' Runner do loop
                    If StrComp(sA(run - n1), item, mCaseSens) <> mComp Then Exit Do
                    run = run - n1
                Loop: cast = lbA + inter1 - run
                If cast <> n0 Then
                    CopyMemByV lpL_1, lpS + (run * n4), cast * n4        ' Grab items that stayed below current that should also be above items that have moved down below current
                    CopyMemByV lpS + (run * n4), lpS + ((lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move up items
                    CopyMemByV lpS + ((lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                End If
            End If '1 2 i 3 4
            If (inter2 <> n0) And (ptr2 - inter2 < lPrettyReverse) Then
                base = lbA + ptr1 + n1
                CopyMemByV lpStr, lpS + (base * n4), n4
                pvt = lbA + ptr1 + inter2
                optimal = pvt + ((ptr2 - inter2) \ n2)
                run = pvt
                Do While run < optimal                                   ' Runner do loop
                    If StrComp(sA(run + n1), item, mCaseSens) <> mComp Then Exit Do
                    run = run + n1
                Loop: cast = run - pvt
                If cast <> n0 Then
                    CopyMemByV lpL_1, lpS + ((pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                    CopyMemByV lpS + ((base + cast) * n4), lpS + (base * n4), inter2 * n4 ' Move up items
                    CopyMemByV lpS + (base * n4), lpL_1, cast * n4       ' Re-assign items into position immediately above current item
        End If: End If: End If

        If (ptr1 > n1) Then
            If (ptr2 > n1) Then cnt = cnt + n1: lbs(cnt) = lbA + ptr1 + n1: ubs(cnt) = ubA
            ubA = lbA + ptr1 - n1
        ElseIf (ptr2 > n1) Then
            lbA = lbA + ptr1 + n1
        Else
            If cnt = n0 Then Exit Do
            lbA = lbs(cnt): ubA = ubs(cnt): cnt = cnt - n1
        End If
    Loop: Erase lbs: Erase ubs: CopyMem ByVal lpStr, 0&, n4
End Sub

' + Stable SwapSort v2 Indexed Version +++++++++++

' This is a stable non-recursive indexed swap sort.

' It is noticably faster than the non-indexed version!
' In fact, it is very nearly as fast as my fastest non-
' stable quicksorts!

' It has the benifit of indexing which allows the source
' array to remain unchanged. This also allows the index
' array to be passed on to other sort processes to be
' further manipulated.

Private Sub strSwapStable2Indexed(sA() As String, idxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    ' This is my indexed stable non-recursive swap sort
    Dim item As String, lpStr As Long, lpS As Long
    Dim Idx As Long, optimal As Long, pvt As Long
    Dim ptr1 As Long, ptr2 As Long, cnt As Long
    Dim base As Long, run As Long, cast As Long
    Dim inter1 As Long, inter2 As Long
    Dim lA_1() As Long, lpL_1 As Long
    Dim lA_2() As Long, lpL_2 As Long
    Dim idxItem As Long, lpI As Long
    Dim lPrettyReverse As Long

    Idx = ubA - lbA ' cnt-1
    ReDim lA_1(n0 To Idx) As Long
    ReDim lA_2(n0 To Idx) As Long
    lpL_1 = VarPtr(lA_1(n0))
    lpL_2 = VarPtr(lA_2(n0))

    pvt = (Idx \ n8) + n16             ' Allow for worst case senario + some
    ReDim lbs(n1 To pvt) As Long       ' Stack to hold pending lower boundries
    ReDim ubs(n1 To pvt) As Long       ' Stack to hold pending upper boundries
    lpStr = VarPtr(item)               ' Cache pointer to the string variable
    lpS = VarPtr(sA(lbA)) - (lbA * n4) ' Cache pointer to the string array

    Do: ptr1 = n0: ptr2 = n0
        pvt = ((ubA - lbA) \ n2) + lbA ' Get pivot index position
        idxItem = idxA(pvt)            ' Grab current value into item
        CopyMemByV lpStr, lpS + (idxItem * n4), n4
        lpI = VarPtr(idxA(lbA))        ' Cache pointer to the index array

        For Idx = lbA To pvt - n1
            If (StrComp(sA(idxA(Idx)), item, mCaseSens) = mComp) Then ' (idx > item)
                lA_2(ptr2) = idxA(Idx) '3
                ptr2 = ptr2 + n1
            Else
                lA_1(ptr1) = idxA(Idx) '1
                ptr1 = ptr1 + n1
            End If
        Next
        inter1 = ptr1: inter2 = ptr2
        For Idx = pvt + n1 To ubA
            If (StrComp(item, sA(idxA(Idx)), mCaseSens) = mComp) Then ' (idx < item)
                lA_1(ptr1) = idxA(Idx) '2
                ptr1 = ptr1 + n1
            Else
                lA_2(ptr2) = idxA(Idx) '4
                ptr2 = ptr2 + n1
            End If
        Next '-Avalanche v2 ©Rd-
        CopyMemByV lpI, lpL_1, ptr1 * n4
        idxA(lbA + ptr1) = idxItem
        CopyMemByV lpI + ((ptr1 + n1) * n4), lpL_2, ptr2 * n4
        lpI = lpI - (lbA * n4)

        If lPrettyReverse <> n0 Then
        ElseIf (ubA - lbA < 250&) Then ' Ignore false indicators
        ElseIf (inter1 = n0) Then
            If (inter2 = ptr2) Then    ' Reverse
                lPrettyReverse = 10000
            ElseIf (ptr1 = n0) Then    ' Pretty
                lPrettyReverse = 50000
            End If
        ElseIf (inter2 = n0) Then      ' Pretty
            If (ptr2 = n0) Then
                lPrettyReverse = 50000
        End If: End If

        If (lPrettyReverse <> n0) Then
            If (ptr1 - inter1 <> n0) And (inter1 < lPrettyReverse) Then  ' Runners dislike super large ranges
                CopyMemByV lpStr, lpS + (idxA(lbA + ptr1 - n1) * n4), n4
                optimal = lbA + (inter1 \ n2)
                run = lbA + inter1
                Do While run > optimal                                   ' Runner do loop
                    If StrComp(sA(idxA(run - n1)), item, mCaseSens) <> mComp Then Exit Do
                    run = run - n1
                Loop: cast = lbA + inter1 - run
                If (cast <> n0) Then
                    CopyMemByV lpL_1, lpI + (run * n4), cast * n4        ' Grab items that stayed below current that should also be above items that have moved down below current
                    CopyMemByV lpI + (run * n4), lpI + ((lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move up items
                    CopyMemByV lpI + ((lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign items into position immediately below current item
                End If
            End If '1 2 i 3 4
            If (inter2 <> n0) And (ptr2 - inter2 < lPrettyReverse) Then
                base = lbA + ptr1 + n1
                CopyMemByV lpStr, lpS + (idxA(base) * n4), n4
                pvt = lbA + ptr1 + inter2
                optimal = pvt + ((ptr2 - inter2) \ n2)
                run = pvt
                Do While run < optimal                                   ' Runner do loop
                    If StrComp(sA(idxA(run + n1)), item, mCaseSens) <> mComp Then Exit Do
                    run = run + n1
                Loop: cast = run - pvt
                If (cast <> n0) Then
                    CopyMemByV lpL_1, lpI + ((pvt + n1) * n4), cast * n4 ' Grab items that stayed above current that should also be below items that have moved up above current
                    CopyMemByV lpI + ((base + cast) * n4), lpI + (base * n4), inter2 * n4 ' Move up items
                    CopyMemByV lpI + (base * n4), lpL_1, cast * n4       ' Re-assign items into position immediately above current item
        End If: End If: End If

        If (ptr1 > n1) Then
            If (ptr2 > n1) Then cnt = cnt + n1: lbs(cnt) = lbA + ptr1 + n1: ubs(cnt) = ubA
            ubA = lbA + ptr1 - n1
        ElseIf (ptr2 > n1) Then
            lbA = lbA + ptr1 + n1
        Else
            If (cnt = n0) Then Exit Do
            lbA = lbs(cnt): ubA = ubs(cnt): cnt = cnt - n1
        End If
    Loop: Erase lbs: Erase ubs: CopyMem ByVal lpStr, 0&, n4
End Sub

' + Pretty Sort ++++++++++++++++++++++++++++++++++++++++

' Sort with binary comparison to seperate upper and lower
' case letters in the order specified by CapsFirst.

' Then sort in the desired direction with case-insensitive
' comparison to group upper and lower case letters together,
' but with a stable sort to preserve the requested caps-first
' or lower-first order.

Sub strPrettySort(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CapsFirst As Boolean = True, Optional ByVal Direction As eDirection = dDefault)
    If Not InitedArray(sA, lbA, ubA) Then Exit Sub
    Dim caseSense As eCompare

    If Direction = dDefault Then Direction = SortDirection
    caseSense = CaseSensitivity

    'False(0) >> dDescending(-1) : True(-1) >> dAscending(1)
    SortCompare = (CapsFirst * -2) - 1
    CaseSensitivity = vbBinaryCompare
    strSwapStable2 sA, lbA, ubA

    SortCompare = Direction
    CaseSensitivity = vbTextCompare
    strSwapStable2 sA, lbA, ubA

    CaseSensitivity = caseSense
End Sub

' + Validate Index Array +++++++++++++++++++++++++++++++++++++++

' This will initialize the passed index array if it is not already.

' This sub-routine requires that the index array be passed either
' prepared for the sort process (see the For loop) or that it be
' uninitialized (or Erased).

' This permits subsequent sorting of the data without interfering
' with the index array if it is already sorted (based on criteria
' that may differ from the current process) and so is not in its
' uninitialized or primary pre-sort state produced by the For loop.

Sub ValidateIndexArray(lIdxA() As Long, ByVal lbA As Long, ByVal ubA As Long)
    Dim bReDim As Boolean, lb As Long, ub As Long, j As Long
    lb = &H80000000: ub = &H7FFFFFFF
    bReDim = Not InitedArray(lIdxA, lb, ub)
    If bReDim = False Then
        bReDim = lbA < lb Or ubA > ub
    End If
    If bReDim Then
        ReDim lIdxA(lbA To ubA) As Long
        For j = lbA To ubA
            lIdxA(j) = j
        Next
    End If
End Sub

' + Inited Array +++++++++++++++++++++++++++++++++++++++++++

' This function determines if the passed array is initialized,
' and if so will return -1.

' It will also optionally indicate whether the array can be
' redimmed - in which case it will return -2.

' If the array is uninitialized (has never been redimmed or
' has been erased) it will return 0 (zero).

Function InitedArray(Arr, lbA As Long, ubA As Long, Optional ByVal bTestRedimable As Boolean) As Long
    ' Thanks to Francesco Balena who solved the Variant
    ' headache, and to Monte Hansen for the ByRef fix
    Dim tSA As SAFEARRAY, tSAB As SABOUNDS, lpSA As Long
    Dim iDataType As Integer, lOffset As Long
    On Error GoTo UnInit
    CopyMem iDataType, Arr, n2                    ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then     ' if a valid string array was passed
        CopyMem lpSA, ByVal VarPtr(Arr) + n8, n4  ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array
        If (iDataType And VT_BYREF) Then          ' see whether the function was passed a Variant that contains an array, rather than directly an array in the former case ptr already points to the SA structure. Thanks to Monte Hansen for this fix
            CopyMem lpSA, ByVal lpSA, n4          ' lpSA is a discripter (pointer) to the safearray structure
        End If
        InitedArray = (lpSA <> n0)
        If InitedArray Then
            CopyMem tSA.cDims, ByVal lpSA, n4
            If bTestRedimable Then ' Return -2 if redimmable
                InitedArray = InitedArray + _
                 ((tSA.fFeatures And FADF_FIXEDSIZE) <> FADF_FIXEDSIZE)
            End If '-©Rd-
            lOffset = n16 + ((tSA.cDims - n1) * n8)
            CopyMem tSAB.cElements, ByVal lpSA + lOffset, n8
            tSAB.lUBound = tSAB.lLBound + tSAB.cElements - n1
            If (lbA < tSAB.lLBound) Then lbA = tSAB.lLBound
            If (ubA > tSAB.lUBound) Then ubA = tSAB.lUBound
    End If: End If
UnInit:
End Function

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++

Function strVerifySort(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CaseSensitive As VbCompareMethod, Optional ByVal Direction As eDirection = dDefault) As Boolean
    On Error GoTo FreakOut
    Dim walk As Long
    For walk = lbA + n1 To ubA
        If StrComp(sA(walk - n1), sA(walk), CaseSensitive) = Direction Then Exit Function
    Next
FreakOut:
    strVerifySort = (walk > ubA)
End Function

Function strVerifyIndexed(sA() As String, lA() As Long, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal CaseSensitive As VbCompareMethod, Optional ByVal Direction As eDirection = dDefault) As Boolean
    On Error GoTo FreakOut
    Dim walk As Long
    For walk = lbA + n1 To ubA
        If StrComp(sA(lA(walk - n1)), sA(lA(walk)), CaseSensitive) = Direction Then Exit Function
    Next
FreakOut:
    strVerifyIndexed = (walk > ubA)
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Rd - crYptic but cRaZy!

