VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuickie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
' A Quicker Quicksort by Ulli
'

Option Explicit
Private Declare Sub PokeLong Lib "kernel32" Alias "RtlMoveMemory" (DestAddress As Long, SourceAddress As Long, Optional ByVal Length As Long = 4)

Private Const Compare_Hex As String = "57 56 8B 7C 24 10 8B 3F 8B 74 24 14 31 C0 39 F7 74 20 39 C7 74 1B 39 C6 74 15 8B 4E FC 3B 4F FC 0F 4F 4F FC D1 E9 41 FC F3 66 A7 74 05 7C 02 48 48 40 5E 5F 8B 54 24 10 89 02 C2 10 00"
Private Compare_Bin()     As Byte

Private Sub Class_Initialize()

    Patch 2, Compare_Hex, Compare_Bin

End Sub

Private Function CompStr(ByVal p1 As Long, ByVal p2 As Long) As Long

    CompStr = 0

    'dummy; will be patched via vTable to:

    'use32
    '        push  edi              ; save edi
    '        push  esi              ; save esi
    '        mov   edi, [esp + 16]  ; get 1st param
    '        mov   edi, [edi]
    '        mov   esi, [esp + 20]  ; get 2nd param
    '        xor   eax, eax         ; clear return value
    '        cmp   edi, esi         ; is same string or both nullstring
    '        je    Equal            ; then equal
    '        cmp   edi, eax         ; is 1st nullstring
    '        je    Less             ; then less
    '        cmp   esi, eax         ; is second nullstring
    '        je    Greater          ; then greater
    '        mov   ecx, [esi - 4]   ; length of 2nd string in bytes
    '        cmp   ecx, [edi - 4]   ; compare with length of 1st string
    '        cmovg ecx, [edi - 4]   ; get longer one for count
    '        shr   ecx,1            ; divide by 2 for unicode
    '        inc   ecx              ; plus 1 for term
    '        cld                    ; forward scan
    '        repe  cmpsw            ; compare strings
    '        je    Equal            ; equal - exit
    '        jl    Less             ; [edi] is less: jump

    'Greater:
    '        dec   eax              ; [edi] is greater: eax=-2; will be incremented
    '        dec   eax

    'Less:
    '        inc   eax              ; eax = 1 or -1

    'Equal:
    '        pop   esi              ; restore esi
    '        pop   edi              ; restore edi
    '        mov   edx, [esp + 16]
    '        mov   [edx], eax
    '        ret   16

End Function

Private Sub DoSort(ByVal pFrom As Long, ByVal pThru As Long)

  Dim pLeft     As Long
  Dim pRite     As Long
  Dim pPivot    As Long

    Do While pFrom < pThru  'we have something to sort (@ least two elements)
        pLeft = pFrom
        pRite = pThru
        PokeLong pPivot, ByVal pLeft 'this is the StrPtr of the pivot element
        Do
            Do Until pRite = pLeft
                If CompStr(pRite, pPivot) < 0 Then 'is less than pivot
                    PokeLong ByVal pLeft, ByVal pRite
                    pLeft = pLeft + 4
                    Exit Do 'loop 
                End If
                pRite = pRite - 4
            Loop
            Do Until pLeft = pRite
                If CompStr(pLeft, pPivot) > 0 Then 'is greater than pivot
                    PokeLong ByVal pRite, ByVal pLeft
                    pRite = pRite - 4
                    Exit Do 'loop 
                End If
                pLeft = pLeft + 4
            Loop
        Loop Until pLeft = pRite
        'now the indexes have met and all bigger items are to the right and all smaller items are left
        PokeLong ByVal pRite, pPivot 'put pivot to where it belongs
        If pLeft - pFrom < pThru - pRite Then 'smaller part 1st to reduce recursion depth
            pLeft = pFrom
            pFrom = pRite + 4
            pRite = pRite - 4
          Else 'NOT pLEFT...
            pRite = pThru
            pThru = pLeft - 4
            pLeft = pLeft + 4
        End If
        If pLeft < pRite Then 'smaller part is not empty...
            DoSort pLeft, pRite '...so sort it
        End If
    Loop

End Sub

Private Sub Patch(ByVal FON As Long, HexCode As String, BinCode() As Byte) 'FON is the Function's Ordinal Number in vTable

  'Convert hex to binary and patch vTable entry for Function CompStr

  Dim sHC()         As String
  Dim vTableAddress As Long
  Dim i             As Long

    sHC = Split(HexCode, " ")
    ReDim BinCode(0 To UBound(sHC))

    'Convert hex to binary
    For i = 0 To UBound(sHC)
        BinCode(i) = Val("&H" & sHC(i))
    Next i

    'Patch vTable entry
    PokeLong vTableAddress, ByVal ObjPtr(Me) 'get vTable address
    PokeLong ByVal vTableAddress + FON * 4 + 28, VarPtr(BinCode(0)) 'patch proper entry in vTable

End Sub

Public Sub SuperQuickie(SortElems() As String, ByVal xFrom As Long, ByVal xThru As Long)

    DoSort VarPtr(SortElems(xFrom)), VarPtr(SortElems(xThru))

End Sub

':) Ulli's VB Code Formatter V2.23.12 (2007-Mrz-25 13:48)  Decl: 5  Code: 128  Total: 133 Lines
':) CommentOnly: 40 (30,1%)  Commented: 12 (9%)  Empty: 27 (20,3%)  Max Logic Depth: 5
