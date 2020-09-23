VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMSDRadixSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************************
'
'    Module:            CMSDRadixSort [CMSDRadixSort.cls]
'    Author:            Ralph Eastwood
'    Email:             tcmreastwood@ntlworld.com
'    Creation:          Tuesday 18, July 2006, 10:14:30 AM
'    Last Modification: Sunday 24, December 2006, 7:58:08 AM
'    Purpose:           MSD Radix Sorts
'    Notes:
'    Update History:    - Created (18/07/2006)
'
'***********************************************************************************************

Option Explicit

Private Type SAFEARRAYHEADER
    cDims       As Integer
    fFeatures   As Integer
    cbElements  As Long
    cLocks      As Long
    pvData      As Long
    cElements   As Long
    lLBound     As Long
End Type

Private Const FADF_AUTO As Long = &H1        '// Array is allocated on the stack.
Private Const FADF_FIXEDSIZE As Long = &H10  '// Array may not be resized or reallocated.

Private Const TA_SIZE As Long = 1000
Private Const TA_THRESHOLD As Long = 30

Private Const AF_SIZE As Long = 1024
Private Const AF_THRESHOLD As Long = 10

Private Const MAX_CHARACTER As Long = 65536

Private Type STRING_STACK
    lpStringArray As Long
    lStringCount As Long
    lStringIndex As Long
End Type

Private m_atStack() As STRING_STACK
Private m_saStringHeader1 As SAFEARRAYHEADER
Private m_aiString1() As Integer
Private m_saStringHeader2 As SAFEARRAYHEADER
Private m_aiString2() As Integer
Private m_saStringPtrArrayHeader As SAFEARRAYHEADER
Private m_alStringPtrArray() As Long
Private m_alPile(0 To MAX_CHARACTER - 1) As Long
Private m_alCount(0 To MAX_CHARACTER - 1) As Long

Private Declare Function GetMem2Long Lib "msvbvm60.dll" Alias "GetMem2" (lpAny As Any, lReturnValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (lpAny As Any, lReturnValue As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60.dll" (lpAny As Any, ByVal lValue As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (dest As Any, src As Any, ByVal cb As Long)
Private Declare Function CompareString Lib "kernel32" Alias "CompareStringW" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long
Private Declare Function SysStringLen Lib "oleaut32" (ByVal lpString As Long) As Long

'***********************************************************************************************
'
'    Procedure: MSD_AF_RadixSort [Public Sub ]
'    Hybrid American flag sort (with stack control), a radix sort
'    algorithm for arrays of character strings by McIlroy, Bostic,
'    and McIlroy.
'
'    Ported to Visual Basic By Ralph Eastwood (tcmreastwood@ntlworld.com)
'
'    Note for VB Version: I've tried my best to comment the sort clearly
'
'***********************************************************************************************
Friend Sub MSD_AF_RadixSort(ByRef A() As String, ByVal lStringCount As Long)

    Dim lStackPointer As Long
    Dim lTmp As Long
    Dim lOldStackPointer As Long
    Dim lBigStackPointer As Long
    Dim lCurrentString As Long, lArrayPointer As Long, lLastString As Long
    Dim i As Long, j As Long
    Dim lCharacterIndex As Long, lCharacter As Long
    Dim lMaxCharacter As Long, lMinCharacter As Long, lNumberOfCharacters As Long
    Dim lCount As Long
    Dim lStackSize As Long
    Dim tTmp As STRING_STACK
    
    lStackSize = AF_SIZE - 1
    ReDim m_atStack(lStackSize)
    
    'Push
    With m_atStack(0)
        .lpStringArray = VarPtr(A(0))
        .lStringCount = lStringCount
        .lStringIndex = 0
    End With
    lStackPointer = 1
    lArrayPointer = 0
    
    Do While lStackPointer 'While the stack is Not empty
        'Pop
        lStackPointer = lStackPointer - 1
        With m_atStack(lStackPointer)
            lArrayPointer = .lpStringArray
            lStringCount = .lStringCount
            lCharacterIndex = .lStringIndex
        End With
        
        m_saStringPtrArrayHeader.pvData = lArrayPointer
        
        If lStringCount < AF_THRESHOLD Then
            'Do a simple insertion sort (faster for less elements) if there's less items
            For i = 1 To lStringCount - 1
                lTmp = m_alStringPtrArray(i)
        
                'Shift elements down until insertion point found
                For j = i - 1 To 0 Step -1
                    
                    'Compare Strings
                    m_saStringHeader1.pvData = m_alStringPtrArray(j) + lCharacterIndex + lCharacterIndex
                    m_saStringHeader2.pvData = lTmp + lCharacterIndex + lCharacterIndex
                    lCharacter = 0
                    Do While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
                        lCharacter = lCharacter + 1
                    Loop
                    'If m_aiString1(lCharacter) = 0 Then lCharacter = lCharacter - 1
                    'if string2 >= string1 then leave the loop
                    If m_aiString2(lCharacter) - m_aiString1(lCharacter) >= 0 Then Exit For
            
                    m_alStringPtrArray(j + 1) = m_alStringPtrArray(j)
                Next j
        
                'Insert
                m_alStringPtrArray(j + 1) = lTmp
                
            Next i
            
        Else

            lLastString = lStringCount
        
            If lNumberOfCharacters = 0 Then
                
                lMinCharacter = MAX_CHARACTER - 1
                lCurrentString = 0
                
                'Get characters from each index and stick it into their respective buckets
                Do While lCurrentString < lLastString
                    
                    'Get a character
                    m_saStringHeader1.pvData = m_alStringPtrArray(lCurrentString)
                    lCurrentString = lCurrentString + 1
                    lCharacter = m_aiString1(lCharacterIndex) And &HFFFF&
                    
                    'Add the character to its bucket
                    m_alCount(lCharacter) = m_alCount(lCharacter) + 1
                    
                    'If there is a character and its not a null
                    If (m_alCount(lCharacter) = 1) And lCharacter Then
                        'Set the minimum
                        If lCharacter < lMinCharacter Then lMinCharacter = lCharacter
                        lNumberOfCharacters = lNumberOfCharacters + 1 'Increment the count of different types of characters

                    End If
                    
                    'Check if we need to resize our stack - to prevent stack overflows
                    If lStackPointer + lNumberOfCharacters > lStackSize Then
                        lStackSize = lStackSize + AF_SIZE
                        ReDim Preserve m_atStack(lStackSize)
                    End If
                Loop
           
                'Logarithmic stack
                lBigStackPointer = lStackPointer
                lOldStackPointer = lStackPointer
                lCount = 2
                
                'Find Places
                lMaxCharacter = 0
                lCurrentString = lArrayPointer + m_alCount(0) * 4 'Current String Pointer
                m_alPile(0) = lCurrentString
                
                'Set the lCharacter to the minimum character, so we can start searching from there
                lCharacter = lMinCharacter
                
                'Go through each character
                Do While lNumberOfCharacters
                    
                    'Search for lLastString existing character
                    Do While m_alCount(lCharacter) = 0
                        lCharacter = lCharacter + 1
                    Loop
                    
                    'If the character exists
                    If m_alCount(lCharacter) > 1 Then
                    
                        If m_alCount(lCharacter) > lCount Then
                            lCount = m_alCount(lCharacter)
                            lBigStackPointer = lStackPointer
                        End If
                                                
                        'Push info for the next character
                        With m_atStack(lStackPointer)
                            .lpStringArray = lCurrentString
                            .lStringCount = m_alCount(lCharacter)
                            .lStringIndex = lCharacterIndex + 1
                        End With
    
                        lStackPointer = lStackPointer + 1
    
                        'Alternative place to put stack resizing
                        'Check if we need to resize our stack - to prevent stack overflows
                        'If lStackPointer > lStackSize Then
                        '    lStackSize = lStackSize + TA_SIZE
                        '    ReDim Preserve m_atStack(lStackSize)
                        'End If
                        
                    End If
                    
                    'Current String Array Pointer (Element in string array pointing to string)
                    lCurrentString = lCurrentString + m_alCount(lCharacter) * 4
                    
                    'Put the current string array ptr to the pile for that character
                    m_alPile(lCharacter) = lCurrentString
                    lMaxCharacter = lCharacter
                    
                    lCharacter = lCharacter + 1
                    lNumberOfCharacters = lNumberOfCharacters - 1
                Loop
        
                'Swap the string element ptrs
                tTmp = m_atStack(lBigStackPointer)
                m_atStack(lBigStackPointer) = m_atStack(lOldStackPointer)
                m_atStack(lOldStackPointer) = tTmp
                
                'Permute Home
                lLastString = lLastString - m_alCount(lMaxCharacter)
                m_alCount(lMaxCharacter) = 0
                
                lCurrentString = 0
                'While its not the last string
                Do While lCurrentString < lLastString
                    
                    m_saStringHeader1.pvData = m_alStringPtrArray(lCurrentString)
                    'Get the character
                    lCharacter = m_aiString1(lCharacterIndex) And &HFFFF&
                    
                    'Decrement the string pointer in the pile i.e. move to the next string
                    m_alPile(lCharacter) = m_alPile(lCharacter) - 4
                    
                    Do While m_alPile(lCharacter) > lArrayPointer + lCurrentString * 4
                        
                        'Using memory access functions - but slightly slower than array access
                        'Left here for clarity - functionally equal to the following 4 lines
                        'GetMem4 ByVal m_alPile(lCharacter), lTmp
                        'PutMem4 ByVal m_alPile(lCharacter), m_saStringHeader1.pvData
                        'm_saStringHeader1.pvData = lTmp
                        
                        'Swap the strings, one stored in m_alPile and currently used
                        m_saStringPtrArrayHeader.pvData = m_alPile(lCharacter)
                        lTmp = m_alStringPtrArray(0)
                        m_alStringPtrArray(0) = m_saStringHeader1.pvData
                        m_saStringHeader1.pvData = lTmp
                        
                        'Get a character
                        lCharacter = m_aiString1(lCharacterIndex) And &HFFFF&
                        m_alPile(lCharacter) = m_alPile(lCharacter) - 4
    
                    Loop
                    
                    m_saStringPtrArrayHeader.pvData = lArrayPointer
                    m_alStringPtrArray(lCurrentString) = m_saStringHeader1.pvData
                    
                    'Increment the string index
                    lCurrentString = lCurrentString + m_alCount(lCharacter)
                    m_alCount(lCharacter) = 0
                Loop
    
            End If
        End If
    Loop
    
End Sub

'***********************************************************************************************
'
'    Procedure: MSD_TA_RadixSort [Public Sub ]
'    Two Array sort (with stack control), a radix sort
'    algorithm for arrays of character strings by McIlroy, Bostic,
'    and McIlroy.
'
'
'    Ported to Visual Basic By Ralph Eastwood (tcmreastwood@ntlworld.com)
'
'***********************************************************************************************
Friend Sub MSD_TA_RadixSort(ByRef A() As String, ByVal lStringCount As Long)

    Dim lStackPointer As Long
    Dim lCurrentString As Long, ai As Long, lTmp As Long, lArrayPointer As Long
    Dim lCharacterIndex As Long, lCharacter As Long, lMinCharacter As Long, lNumberOfCharacters As Long
    Dim ta() As Long
    Dim i As Long, j As Long, k As Long
    Dim lStringPtr As Long
    Dim lStackSize As Long
    
    lStackSize = TA_SIZE - 1
    ReDim ta(lStringCount - 1)
    ReDim m_atStack(lStackSize)
    
    'Push
    With m_atStack(0)
        .lpStringArray = VarPtr(A(0))
        .lStringCount = lStringCount
        .lStringIndex = 0
    End With
    lStackPointer = 1
    lArrayPointer = 0

    Do While lStackPointer    'While the stack is Not empty
    
        'Pop
        lStackPointer = lStackPointer - 1
        With m_atStack(lStackPointer)
            lArrayPointer = .lpStringArray
            lStringCount = .lStringCount
            lCharacterIndex = .lStringIndex
        End With

        m_saStringPtrArrayHeader.pvData = lArrayPointer

        
        If lStringCount < TA_THRESHOLD Then
            'Do a simple insertion sort (faster for less elements) if there's less items
            For i = 1 To lStringCount - 1
                lTmp = m_alStringPtrArray(i)
        
                'Shift elements down until insertion point found
                For j = i - 1 To 0 Step -1
                    
                    'Compare Strings
                    m_saStringHeader1.pvData = m_alStringPtrArray(j) + lCharacterIndex + lCharacterIndex
                    m_saStringHeader2.pvData = lTmp + lCharacterIndex + lCharacterIndex
                    lCharacter = 0
                    While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
                        lCharacter = lCharacter + 1
                    Wend
                    'If m_aiString1(lCharacter) = 0 Then lCharacter = lCharacter - 1
                    If m_aiString2(lCharacter) - m_aiString1(lCharacter) >= 0 Then Exit For
            
                    m_alStringPtrArray(j + 1) = m_alStringPtrArray(j)
                Next j
        
                'Insert
                m_alStringPtrArray(j + 1) = lTmp
                
            Next i
            
        Else

            lMinCharacter = MAX_CHARACTER - 1
            lCurrentString = lStringCount

            Do While lCurrentString > 0
                lCurrentString = lCurrentString - 1
                m_saStringHeader1.pvData = m_alStringPtrArray(lCurrentString)
                lCharacter = m_aiString1(lCharacterIndex) And &HFFFF&
                m_alCount(lCharacter) = m_alCount(lCharacter) + 1
                If m_alCount(lCharacter) = 1 And lCharacter > 0 Then
                    If lCharacter < lMinCharacter Then lMinCharacter = lCharacter
                    lNumberOfCharacters = lNumberOfCharacters + 1
                End If
                If lStackPointer + lNumberOfCharacters > lStackSize Then
                    lStackSize = lStackSize + AF_SIZE
                    ReDim Preserve m_atStack(lStackSize)
                End If
            Loop
            lCurrentString = lArrayPointer + m_alCount(0) * 4
            m_alPile(0) = lCurrentString
            m_alCount(0) = 0

            lCharacter = lMinCharacter
            Do While lNumberOfCharacters > 0

                While m_alCount(lCharacter) = 0
                    lCharacter = lCharacter + 1
                Wend

                If m_alCount(lCharacter) > 1 Then
                    With m_atStack(lStackPointer)
                        .lpStringArray = lCurrentString
                        .lStringCount = m_alCount(lCharacter)
                        .lStringIndex = lCharacterIndex + 1
                    End With
                    lStackPointer = lStackPointer + 1
                    
                    'Allocate more memory if there is a stack overflow
                    'If lStackPointer > lStackSize Then
                    '    lStackSize = lStackSize + TA_SIZE
                    '    ReDim Preserve m_atStack(lStackSize)
                    'End If
                    
                End If

                lCurrentString = lCurrentString + m_alCount(lCharacter) * 4
                m_alPile(lCharacter) = lCurrentString
                m_alCount(lCharacter) = 0

                lCharacter = lCharacter + 1
                lNumberOfCharacters = lNumberOfCharacters - 1
            Loop

            lCurrentString = lStringCount
            ai = lStringCount

            Do While lCurrentString > 0
                lCurrentString = lCurrentString - 1
                ai = ai - 1
                ta(lCurrentString) = m_alStringPtrArray(ai)
            Loop
            lCurrentString = lStringCount - 1
            Do While lCurrentString >= 0
                m_saStringHeader1.pvData = ta(lCurrentString)
                lTmp = m_aiString1(lCharacterIndex) And &HFFFF&
                m_alPile(lTmp) = m_alPile(lTmp) - 4
                
                'PutMem4 ByVal m_alPile(lTmp), ta(lCurrentString)
                m_saStringPtrArrayHeader.pvData = m_alPile(lTmp)
                m_alStringPtrArray(0) = ta(lCurrentString)
                lCurrentString = lCurrentString - 1
            Loop
        End If
        
    Loop

End Sub
'***********************************************************************************************
'
'    Procedure: MSD_AF_RadixSort_DSC [Public Sub ]
'    Hybrid American flag sort (with stack control), a radix sort
'    algorithm for arrays of character strings by McIlroy, Bostic,
'    and McIlroy. (Descending)
'
'    Ported to Visual Basic By Ralph Eastwood (tcmreastwood@ntlworld.com)
'
'    Note for VB Version: I've tried my best to comment the sort clearly
'
'***********************************************************************************************
Friend Sub MSD_AF_RadixSort_DSC(ByRef A() As String, ByVal lStringCount As Long)

    Dim lStackPointer As Long
    Dim lTmp As Long
    Dim lOldStackPointer As Long
    Dim lBigStackPointer As Long
    Dim lCurrentString As Long, lArrayPointer As Long, lLastString As Long
    Dim i As Long, j As Long
    Dim lCharacterIndex As Long, lCharacter As Long
    Dim lMaxCharacter As Long, lMinCharacter As Long, lNumberOfCharacters As Long
    Dim lCount As Long
    Dim lStackSize As Long
    Dim tTmp As STRING_STACK
    
    lStackSize = AF_SIZE - 1
    ReDim m_atStack(lStackSize)
    
    'Push
    With m_atStack(0)
        .lpStringArray = VarPtr(A(0))
        .lStringCount = lStringCount
        .lStringIndex = 0
    End With
    lStackPointer = 1
    lArrayPointer = 0
    
    Do While lStackPointer 'While the stack is Not empty
        'Pop
        lStackPointer = lStackPointer - 1
        With m_atStack(lStackPointer)
            lArrayPointer = .lpStringArray
            lStringCount = .lStringCount
            lCharacterIndex = .lStringIndex
        End With
        
        m_saStringPtrArrayHeader.pvData = lArrayPointer
        
        If lStringCount < AF_THRESHOLD Then
            'Do a simple insertion sort (faster for less elements) if there's less items
            For i = 1 To lStringCount - 1
                lTmp = m_alStringPtrArray(i)
        
                'Shift elements down until insertion point found
                For j = i - 1 To 0 Step -1
                    
                    'Compare Strings
                    m_saStringHeader1.pvData = m_alStringPtrArray(j) + lCharacterIndex + lCharacterIndex
                    m_saStringHeader2.pvData = lTmp + lCharacterIndex + lCharacterIndex
                    lCharacter = 0
                    Do While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
                        lCharacter = lCharacter + 1
                    Loop
                    'If m_aiString1(lCharacter) = 0 Then lCharacter = lCharacter - 1
                    If m_aiString2(lCharacter) - m_aiString1(lCharacter) <= 0 Then Exit For
            
                    m_alStringPtrArray(j + 1) = m_alStringPtrArray(j)
                Next j
        
                'Insert
                m_alStringPtrArray(j + 1) = lTmp
                
            Next i
            
        Else

            lLastString = lStringCount
        
            If lNumberOfCharacters = 0 Then
                
                lMaxCharacter = MAX_CHARACTER - 1
                lCurrentString = 0
                
                'Get characters from each index and stick it into their respective buckets
                Do While lCurrentString < lLastString
                    
                    'Get a character
                    m_saStringHeader1.pvData = m_alStringPtrArray(lCurrentString)
                    lCurrentString = lCurrentString + 1
                    lCharacter = m_aiString1(lCharacterIndex) And &HFFFF&
                    
                    'Add the character to its bucket
                    m_alCount(lCharacter) = m_alCount(lCharacter) + 1
                    
                    'If there is a character and its not a null
                    If (m_alCount(lCharacter) = 1) And lCharacter Then
                        'Set the minimum
                        If lCharacter > lMinCharacter Then lMinCharacter = lCharacter
                        lNumberOfCharacters = lNumberOfCharacters + 1 'Increment the count of different types of characters

                    End If
                    
                    'Check if we need to resize our stack - to prevent stack overflows
                    If lStackPointer + lNumberOfCharacters > lStackSize Then
                        lStackSize = lStackSize + AF_SIZE
                        ReDim Preserve m_atStack(lStackSize)
                    End If
                Loop
           
                'Logarithmic stack
                lBigStackPointer = lStackPointer
                lOldStackPointer = lStackPointer
                lCount = 2
                
                'Find Places
                lMinCharacter = MAX_CHARACTER - 1
                lCurrentString = lArrayPointer + m_alCount(0) * 4 'Current String Pointer
                m_alPile(0) = lCurrentString
                
                'Set the lCharacter to the maximum character, so we can start searching from there
                lCharacter = lMaxCharacter
                
                'Go through each character
                Do While lNumberOfCharacters
                    
                    'Search for lLastString existing character
                    Do While m_alCount(lCharacter) = 0
                        lCharacter = lCharacter - 1
                    Loop
                    
                    'If the character exists
                    If m_alCount(lCharacter) > 1 Then
                    
                        If m_alCount(lCharacter) > lCount Then
                            lCount = m_alCount(lCharacter)
                            lBigStackPointer = lStackPointer
                        End If
                                                
                        'Push info for the next character
                        With m_atStack(lStackPointer)
                            .lpStringArray = lCurrentString
                            .lStringCount = m_alCount(lCharacter)
                            .lStringIndex = lCharacterIndex + 1
                        End With
    
                        lStackPointer = lStackPointer + 1
    
                        'Alternative place to put stack resizing
                        'Check if we need to resize our stack - to prevent stack overflows
                        'If lStackPointer > lStackSize Then
                        '    lStackSize = lStackSize + TA_SIZE
                        '    ReDim Preserve m_atStack(lStackSize)
                        'End If
                        
                    End If
                    
                    'Current String Array Pointer (Element in string array pointing to string)
                    lCurrentString = lCurrentString + m_alCount(lCharacter) * 4
                    
                    'Put the current string array ptr to the pile for that character
                    m_alPile(lCharacter) = lCurrentString
                    lMinCharacter = lCharacter
                    
                    lCharacter = lCharacter - 1
                    lNumberOfCharacters = lNumberOfCharacters - 1
                Loop
        
                'Swap the string element ptrs
                tTmp = m_atStack(lBigStackPointer)
                m_atStack(lBigStackPointer) = m_atStack(lOldStackPointer)
                m_atStack(lOldStackPointer) = tTmp
                
                'Permute Home
                lLastString = lLastString - m_alCount(lMinCharacter)
                m_alCount(lMinCharacter) = 0
                
                lCurrentString = 0
                'While its not the last string
                Do While lCurrentString < lLastString
                    
                    m_saStringHeader1.pvData = m_alStringPtrArray(lCurrentString)
                    'Get the character
                    lCharacter = m_aiString1(lCharacterIndex) And &HFFFF&
                    
                    'Decrement the string pointer in the pile i.e. move to the next string
                    m_alPile(lCharacter) = m_alPile(lCharacter) - 4
                    
                    Do While m_alPile(lCharacter) > lArrayPointer + lCurrentString * 4
                        
                        'Using memory access functions - but slightly slower than array access
                        'Left here for clarity - functionally equal to the following 4 lines
                        'GetMem4 ByVal m_alPile(lCharacter), lTmp
                        'PutMem4 ByVal m_alPile(lCharacter), m_saStringHeader1.pvData
                        'm_saStringHeader1.pvData = lTmp
                        
                        'Swap the strings, one stored in m_alPile and currently used
                        m_saStringPtrArrayHeader.pvData = m_alPile(lCharacter)
                        lTmp = m_alStringPtrArray(0)
                        m_alStringPtrArray(0) = m_saStringHeader1.pvData
                        m_saStringHeader1.pvData = lTmp
                        
                        'Get a character
                        lCharacter = m_aiString1(lCharacterIndex) And &HFFFF&
                        m_alPile(lCharacter) = m_alPile(lCharacter) - 4
    
                    Loop
                    
                    m_saStringPtrArrayHeader.pvData = lArrayPointer
                    m_alStringPtrArray(lCurrentString) = m_saStringHeader1.pvData
                    
                    'Increment the string index
                    lCurrentString = lCurrentString + m_alCount(lCharacter)
                    m_alCount(lCharacter) = 0
                Loop
    
            End If
        End If
    Loop
    
End Sub
'***********************************************************************************************
'
'    Procedure: MSD_TA_RadixSort_DSC [Public Sub ]
'    Two Array sort (with stack control), a radix sort
'    algorithm for arrays of character strings by McIlroy, Bostic,
'    and McIlroy.
'
'
'    Ported to Visual Basic By Ralph Eastwood (tcmreastwood@ntlworld.com)
'
'***********************************************************************************************
Friend Sub MSD_TA_RadixSort_DSC(ByRef A() As String, ByVal lStringCount As Long)

    Dim lStackPointer As Long
    Dim lCurrentString As Long, ai As Long, lTmp As Long, lArrayPointer As Long
    Dim lCharacterIndex As Long, lCharacter As Long, lMaxCharacter As Long, lNumberOfCharacters As Long
    Dim ta() As Long
    Dim i As Long, j As Long, k As Long
    Dim lStringPtr As Long
    Dim lStackSize As Long
    
    lStackSize = TA_SIZE - 1
    ReDim ta(lStringCount - 1)
    ReDim m_atStack(lStackSize)
    
    'Push
    With m_atStack(0)
        .lpStringArray = VarPtr(A(0))
        .lStringCount = lStringCount
        .lStringIndex = 0
    End With
    lStackPointer = 1
    lArrayPointer = 0

    Do While lStackPointer    'While the stack is Not empty
    
        'Pop
        lStackPointer = lStackPointer - 1
        With m_atStack(lStackPointer)
            lArrayPointer = .lpStringArray
            lStringCount = .lStringCount
            lCharacterIndex = .lStringIndex
        End With

        m_saStringPtrArrayHeader.pvData = lArrayPointer

        
        If lStringCount < TA_THRESHOLD Then
            'Do a simple insertion sort (faster for less elements) if there's less items
            For i = 1 To lStringCount - 1
                lTmp = m_alStringPtrArray(i)
        
                'Shift elements down until insertion point found
                For j = i - 1 To 0 Step -1
                    
                    'Compare Strings
                    m_saStringHeader1.pvData = m_alStringPtrArray(j) + lCharacterIndex + lCharacterIndex
                    m_saStringHeader2.pvData = lTmp + lCharacterIndex + lCharacterIndex
                    lCharacter = 0
                    Do While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
                        lCharacter = lCharacter + 1
                    Loop
                    'If m_aiString1(lCharacter) = 0 Then lCharacter = lCharacter - 1
                    If m_aiString2(lCharacter) - m_aiString1(lCharacter) <= 0 Then Exit For
            
                    m_alStringPtrArray(j + 1) = m_alStringPtrArray(j)
                Next j
        
                'Insert
                m_alStringPtrArray(j + 1) = lTmp
                
            Next i
            
        Else

            lMaxCharacter = 0
            lCurrentString = lStringCount

            Do While lCurrentString > 0
                lCurrentString = lCurrentString - 1
                m_saStringHeader1.pvData = m_alStringPtrArray(lCurrentString)
                lCharacter = m_aiString1(lCharacterIndex) And &HFFFF&
                m_alCount(lCharacter) = m_alCount(lCharacter) + 1
                If m_alCount(lCharacter) = 1 And lCharacter > 0 Then
                    If lCharacter > lMaxCharacter Then lMaxCharacter = lCharacter
                    lNumberOfCharacters = lNumberOfCharacters + 1
                End If
                If lStackPointer + lNumberOfCharacters > lStackSize Then
                    lStackSize = lStackSize + AF_SIZE
                    ReDim Preserve m_atStack(lStackSize)
                End If
            Loop
            lCurrentString = lArrayPointer + m_alCount(0) * 4
            m_alPile(0) = lCurrentString
            m_alCount(0) = 0

            lCharacter = lMaxCharacter
            Do While lNumberOfCharacters > 0

                Do While m_alCount(lCharacter) = 0
                    lCharacter = lCharacter - 1
                Loop

                If m_alCount(lCharacter) > 1 Then
                    With m_atStack(lStackPointer)
                        .lpStringArray = lCurrentString
                        .lStringCount = m_alCount(lCharacter)
                        .lStringIndex = lCharacterIndex + 1
                    End With
                    lStackPointer = lStackPointer + 1
                    
                    'Allocate more memory if there is a stack overflow
                    'If lStackPointer > lStackSize Then
                    '    lStackSize = lStackSize + TA_SIZE
                    '    ReDim Preserve m_atStack(lStackSize)
                    'End If
                    
                End If

                lCurrentString = lCurrentString + m_alCount(lCharacter) * 4
                m_alPile(lCharacter) = lCurrentString
                m_alCount(lCharacter) = 0

                lCharacter = lCharacter + 1
                lNumberOfCharacters = lNumberOfCharacters - 1
            Loop

            lCurrentString = lStringCount
            ai = lStringCount

            Do While lCurrentString > 0
                lCurrentString = lCurrentString - 1
                ai = ai - 1
                ta(lCurrentString) = m_alStringPtrArray(ai)
            Loop
            lCurrentString = lStringCount - 1
            Do While lCurrentString >= 0
                m_saStringHeader1.pvData = ta(lCurrentString)
                lTmp = m_aiString1(lCharacterIndex) And &HFFFF&
                m_alPile(lTmp) = m_alPile(lTmp) - 4
                
                'PutMem4 ByVal m_alPile(lTmp), ta(lCurrentString)
                m_saStringPtrArrayHeader.pvData = m_alPile(lTmp)
                m_alStringPtrArray(0) = ta(lCurrentString)
                lCurrentString = lCurrentString - 1
            Loop
        End If
        
    Loop

End Sub


'***********************************************************************************************
'
'    Procedure:     Class_Initialize [Private Sub ]
'    Purpose:       Initialise the SafeArray Headers
'    Called From:
'    Parameters:    None
'    Returns:       None
'
'***********************************************************************************************

Private Sub Class_Initialize()

    With m_saStringHeader1
        .cDims = 1
        .cbElements = 2
        .lLBound = 0
        .cElements = &H7FFFFFFF
    End With
    PutMem4 ByVal ArrPtr(m_aiString1), ByVal VarPtr(m_saStringHeader1)

    With m_saStringHeader2
        .cDims = 1
        .cbElements = 2
        .lLBound = 0
        .cElements = &H7FFFFFFF
    End With
    PutMem4 ByVal ArrPtr(m_aiString2), ByVal VarPtr(m_saStringHeader2)


    With m_saStringPtrArrayHeader
        .cDims = 1
        .cbElements = 4
        .lLBound = 0
        .cElements = &H7FFFFFFF
    End With
    PutMem4 ByVal ArrPtr(m_alStringPtrArray), ByVal VarPtr(m_saStringPtrArrayHeader)

End Sub

'***********************************************************************************************
'
'    Procedure:     Class_Terminate [Private Sub ]
'    Purpose:       Clear up the SafeArray Headers
'    Called From:
'    Parameters:    None
'    Returns:       None
'
'***********************************************************************************************

Private Sub Class_Terminate()
    
    'Clear up the array pointers so vb doesn't try and free memory twice
    PutMem4 ByVal ArrPtr(m_aiString1), ByVal 0&
    PutMem4 ByVal ArrPtr(m_aiString2), ByVal 0&
    PutMem4 ByVal ArrPtr(m_alStringPtrArray), ByVal 0&
    
End Sub
