VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTriQuickSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************************************************************************
'
'    Module:            CTriQuickSort [CTriQuickSort.cls]
'    Author(s) info:
'     By:         Philippe Lord // Marton
'     Email:      StromgaldMarton@Hotmail.com
'     ICQ:        12181387'
'    Modified and optimised by Ralph Eastwood
'
'***********************************************************************************************

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

Private Const INITIAL_STACK = 128
Private Const RESIZE_STACK = 128
Private Const THRESHOLD_QSR = 16
Private Const THRESHOLD_QSNR = 16

Private m_lLBStack() As Long
Private m_lUBStack() As Long
Private m_lStackSize As Long

Private m_saStringHeader1 As SAFEARRAYHEADER
Private m_aiString1() As Integer
Private m_saStringHeader2 As SAFEARRAYHEADER
Private m_aiString2() As Integer
Private m_saStringPtrArrayHeader As SAFEARRAYHEADER
Private m_alStringPtrArray() As Long

Private Declare Function GetMem4 Lib "msvbvm60.dll" (lpAny As Any, lReturnValue As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60.dll" (lpAny As Any, ByVal lValue As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (dest As Any, src As Any, ByVal cb As Long)

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

Public Sub TriQuickSortString(ByRef sArray() As String, ByVal lLowerBound As Long, ByVal lUpperBound As Long)
   'Dim i       As Long
   'Dim j       As Long
      
   ' Get the String array ptr
   m_saStringPtrArrayHeader.pvData = VarPtr(sArray(0))
    
   ' *NOTE*  the value 4 is VERY important here !!!
   ' DO NOT CHANGE 4 FOR A LOWER VALUE !!!
   TriQuickSortString2 sArray, THRESHOLD_QSR, lLowerBound, lUpperBound
   InsertionSortString sArray, lLowerBound, lUpperBound
   
End Sub
Private Sub TriQuickSortString2(ByRef sArray() As String, ByVal iSplit As Long, ByVal lLowerBound As Long, ByVal lUpperBound As Long)
   Dim i     As Long
   Dim j     As Long
   Dim sTemp As String
   Dim lTemp As Long
   Dim lCharacter As Long
   Dim lPtr As Long
   
   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.

   
   If (lUpperBound - lLowerBound) > iSplit Then
      i = (lUpperBound + lLowerBound) \ 2
      
      'Compare Strings
      m_saStringHeader1.pvData = m_alStringPtrArray(lLowerBound)
      m_saStringHeader2.pvData = m_alStringPtrArray(i)
      lCharacter = 0
      While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
         lCharacter = lCharacter + 1
      Wend
      If m_aiString1(lCharacter) > m_aiString2(lCharacter) Then
      
      'If sArray(lLowerBound) > sArray(i) Then
         'SwapStrings sArray(lLowerBound), sArray(i)
         lTemp = m_alStringPtrArray(i)
         m_alStringPtrArray(i) = m_alStringPtrArray(lLowerBound)
         m_alStringPtrArray(lLowerBound) = lTemp
      End If
        
      m_saStringHeader1.pvData = m_alStringPtrArray(lLowerBound)
      m_saStringHeader2.pvData = m_alStringPtrArray(lUpperBound)
      lCharacter = 0
      While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
         lCharacter = lCharacter + 1
      Wend
      'If sArray(lLowerBound) > sArray(lUpperBound) Then
      If m_aiString1(lCharacter) > m_aiString2(lCharacter) Then
         'SwapStrings sArray(lLowerBound), sArray(lUpperBound)
         lTemp = m_alStringPtrArray(lUpperBound)
         m_alStringPtrArray(lUpperBound) = m_alStringPtrArray(lLowerBound)
         m_alStringPtrArray(lLowerBound) = lTemp
      End If
      
      m_saStringHeader1.pvData = m_alStringPtrArray(i)
      m_saStringHeader2.pvData = m_alStringPtrArray(lUpperBound)
      lCharacter = 0
      While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
         lCharacter = lCharacter + 1
      Wend
      'If sArray(i) > sArray(lUpperBound) Then
      If m_aiString1(lCharacter) > m_aiString2(lCharacter) Then
         'SwapStrings sArray(i), sArray(lUpperBound)
         lTemp = m_alStringPtrArray(lUpperBound)
         m_alStringPtrArray(lUpperBound) = m_alStringPtrArray(i)
         m_alStringPtrArray(i) = lTemp
      End If
      
      j = lUpperBound - 1
      'SwapStrings sArray(i), sArray(j)
      lTemp = m_alStringPtrArray(j)
      m_alStringPtrArray(j) = m_alStringPtrArray(i)
      m_alStringPtrArray(i) = lTemp
      i = lLowerBound
      'RtlMoveMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(j)), 4 ' sTemp = sArray(j)
      'PutMem4 ByVal VarPtr(sTemp), m_alStringPtrArray(j)
      m_saStringHeader2.pvData = m_alStringPtrArray(j)
      
      Do

         'Do
         '   i = i + 1
         'Loop While sArray(i) < sTemp
         Do
            i = i + 1
            m_saStringHeader1.pvData = m_alStringPtrArray(i)
            lCharacter = 0
            Do While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
               lCharacter = lCharacter + 1
            Loop
         Loop While m_aiString1(lCharacter) < m_aiString2(lCharacter)

         'Do
         '   j = j - 1
         'Loop While sArray(j) > sTemp
         
         Do
            j = j - 1
            m_saStringHeader1.pvData = m_alStringPtrArray(j)
            lCharacter = 0
            Do While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
               lCharacter = lCharacter + 1
            Loop
         Loop While m_aiString1(lCharacter) > m_aiString2(lCharacter)
         
         If j < i Then Exit Do
         'SwapStrings sArray(i), sArray(j)
         lTemp = m_alStringPtrArray(j)
         m_alStringPtrArray(j) = m_alStringPtrArray(i)
         m_alStringPtrArray(i) = lTemp
      Loop
      
      'SwapStrings sArray(i), sArray(lUpperBound - 1)
      lTemp = m_alStringPtrArray(lUpperBound - 1)
      m_alStringPtrArray(lUpperBound - 1) = m_alStringPtrArray(i)
      m_alStringPtrArray(i) = lTemp
         
      TriQuickSortString2 sArray, iSplit, lLowerBound, j
      TriQuickSortString2 sArray, iSplit, i + 1, lUpperBound
   End If
    
   ' clear temp var (sTemp)
   'PutMem4 ByVal VarPtr(sTemp), 0
End Sub

Private Sub InsertionSortString(ByRef sArray() As String, ByVal lLowerBound As Long, ByVal lUpperBound As Long)
   Dim i     As Long
   Dim j     As Long
   'Dim sTemp As String
   Dim lTemp As Long
   Dim lCharacter As Long
   
   ' *NOTE* no checks are made in this function because it is used internally.
   ' Validity checks are made in the public function that calls this one.
   ' Get the String array ptr
   m_saStringPtrArrayHeader.pvData = VarPtr(sArray(lLowerBound))

   For i = lLowerBound + 1 To lUpperBound
      'RtlMoveMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(i)), 4 ' sTemp = sArray(i)
      lTemp = m_alStringPtrArray(i)
      
      j = i
      
      Do While j > lLowerBound
         m_saStringHeader1.pvData = m_alStringPtrArray(j - 1)
         m_saStringHeader2.pvData = lTemp
         lCharacter = 0
         Do While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
            lCharacter = lCharacter + 1
         Loop
         If m_aiString1(lCharacter) = 0 Then lCharacter = lCharacter - 1
         'If sArray(j - 1) <= sTemp Then Exit Do
         If m_aiString1(lCharacter) <= m_aiString2(lCharacter) Then Exit Do

         'RtlMoveMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sArray(j - 1)), 4 ' sArray(j) = sArray(j - 1)
         m_alStringPtrArray(j) = m_alStringPtrArray(j - 1)
         j = j - 1
      Loop
      
      'RtlMoveMemory ByVal VarPtr(sArray(j)), ByVal VarPtr(sTemp), 4
      m_alStringPtrArray(j) = lTemp
      ' sArray(j) = sTemp
   Next i
   
   ' clear temp var (sTemp)
   'i = 0
   'RtlMoveMemory ByVal VarPtr(sTemp), ByVal VarPtr(i), 4
End Sub

Public Sub InitTriQSortStack(Optional ByVal lStackSize As Long = INITIAL_STACK)
    m_lStackSize = lStackSize - 1
    ReDim Preserve m_lLBStack(lStackSize - 1)
    ReDim Preserve m_lUBStack(lStackSize - 1)
End Sub


Public Sub TriQSortString(ByRef sArray() As String, ByVal lLowerBound As Long, ByVal lUpperBound As Long)
    Dim i     As Long
    Dim j     As Long
    Dim sTemp As String
    Dim lTemp As Long
    Dim lCharacter As Long
    Dim lPtr As Long
    Dim lStackIndex As Long
    Dim lOldUB As Long
    Dim lOldLB As Long
    
    ' Store the original bounds
    lOldUB = lUpperBound
    lOldLB = lLowerBound
    
    ' Get the String array ptr
    m_saStringPtrArrayHeader.pvData = VarPtr(sArray(lLowerBound))
     
    ' Set the bounds for the stack
    m_lLBStack(0) = lLowerBound
    m_lUBStack(0) = lUpperBound
    lStackIndex = 0
     
    ' If the stack has not been initialised, do so
    If m_lStackSize = 0 Then
        m_lStackSize = INITIAL_STACK - 1
        ReDim Preserve m_lLBStack(m_lStackSize)
        ReDim Preserve m_lUBStack(m_lStackSize)
    End If
 
    Do While lStackIndex >= 0
        If (lUpperBound - lLowerBound) > THRESHOLD_QSNR Then
            
            i = (lUpperBound + lLowerBound) \ 2
            
            'Compare Strings
            m_saStringHeader1.pvData = m_alStringPtrArray(lLowerBound)
            m_saStringHeader2.pvData = m_alStringPtrArray(i)
            lCharacter = 0
            While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
               lCharacter = lCharacter + 1
            Wend
            If m_aiString1(lCharacter) > m_aiString2(lCharacter) Then
            
            'If sArray(lLowerBound) > sArray(i) Then
               'SwapStrings sArray(lLowerBound), sArray(i)
               lTemp = m_alStringPtrArray(i)
               m_alStringPtrArray(i) = m_alStringPtrArray(lLowerBound)
               m_alStringPtrArray(lLowerBound) = lTemp
            End If
              
            m_saStringHeader1.pvData = m_alStringPtrArray(lLowerBound)
            m_saStringHeader2.pvData = m_alStringPtrArray(lUpperBound)
            lCharacter = 0
            While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
               lCharacter = lCharacter + 1
            Wend
            'If sArray(lLowerBound) > sArray(lUpperBound) Then
            If m_aiString1(lCharacter) > m_aiString2(lCharacter) Then
               'SwapStrings sArray(lLowerBound), sArray(lUpperBound)
               lTemp = m_alStringPtrArray(lUpperBound)
               m_alStringPtrArray(lUpperBound) = m_alStringPtrArray(lLowerBound)
               m_alStringPtrArray(lLowerBound) = lTemp
            End If
            
            m_saStringHeader1.pvData = m_alStringPtrArray(i)
            m_saStringHeader2.pvData = m_alStringPtrArray(lUpperBound)
            lCharacter = 0
            While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
               lCharacter = lCharacter + 1
            Wend
            'If sArray(i) > sArray(lUpperBound) Then
            If m_aiString1(lCharacter) > m_aiString2(lCharacter) Then
               'SwapStrings sArray(i), sArray(lUpperBound)
               lTemp = m_alStringPtrArray(lUpperBound)
               m_alStringPtrArray(lUpperBound) = m_alStringPtrArray(i)
               m_alStringPtrArray(i) = lTemp
            End If
            
            j = lUpperBound - 1
            'SwapStrings sArray(i), sArray(j)
            lTemp = m_alStringPtrArray(j)
            m_alStringPtrArray(j) = m_alStringPtrArray(i)
            m_alStringPtrArray(i) = lTemp
            i = lLowerBound
            'RtlMoveMemory ByVal VarPtr(sTemp), ByVal VarPtr(sArray(j)), 4 ' sTemp = sArray(j)
            'PutMem4 ByVal VarPtr(sTemp), m_alStringPtrArray(j)
            lPtr = m_alStringPtrArray(j)
            
            Do
               m_saStringHeader2.pvData = lPtr
            
               'Do
               '   i = i + 1
               'Loop While sArray(i) < sTemp
               Do
                  i = i + 1
                  m_saStringHeader1.pvData = m_alStringPtrArray(i)
                  lCharacter = 0
                  Do While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
                     lCharacter = lCharacter + 1
                  Loop
               Loop While m_aiString1(lCharacter) < m_aiString2(lCharacter)
            
               'Do
               '   j = j - 1
               'Loop While sArray(j) > sTemp
               
               Do
                  j = j - 1
                  m_saStringHeader1.pvData = m_alStringPtrArray(j)
                  lCharacter = 0
                  Do While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
                     lCharacter = lCharacter + 1
                  Loop
               Loop While m_aiString1(lCharacter) > m_aiString2(lCharacter)
               
               If j < i Then Exit Do
               'SwapStrings sArray(i), sArray(j)
               lTemp = m_alStringPtrArray(j)
               m_alStringPtrArray(j) = m_alStringPtrArray(i)
               m_alStringPtrArray(i) = lTemp
            Loop
            
            'SwapStrings sArray(i), sArray(lUpperBound - 1)
            lTemp = m_alStringPtrArray(lUpperBound - 1)
            m_alStringPtrArray(lUpperBound - 1) = m_alStringPtrArray(i)
            m_alStringPtrArray(i) = lTemp
               
            'TriQuickSortString2 sArray, iSplit, lLowerBound, j
            'TriQuickSortString2 sArray, iSplit, i + 1, lUpperBound
            
            If (j - lLowerBound) > THRESHOLD_QSNR Then
                m_lLBStack(lStackIndex) = lLowerBound
                m_lUBStack(lStackIndex) = j
                lStackIndex = lStackIndex + 1
                            
                ' Resize the stack if it overflows
                If lStackIndex >= m_lStackSize Then
                    m_lStackSize = m_lStackSize + RESIZE_STACK
                    ReDim Preserve m_lLBStack(m_lStackSize)
                    ReDim Preserve m_lUBStack(m_lStackSize)
                End If
            End If

            If (lUpperBound - i + 1) > THRESHOLD_QSNR Then
                m_lLBStack(lStackIndex) = i + 1
                m_lUBStack(lStackIndex) = lUpperBound
                lStackIndex = lStackIndex + 1
                
                ' Resize the stack if it overflows
                If lStackIndex >= m_lStackSize Then
                    m_lStackSize = m_lStackSize + RESIZE_STACK
                    ReDim Preserve m_lLBStack(m_lStackSize)
                    ReDim Preserve m_lUBStack(m_lStackSize)
                End If
            End If
        End If
        
        If lStackIndex Then
            lStackIndex = lStackIndex - 1
            
            lLowerBound = m_lLBStack(lStackIndex)
            lUpperBound = m_lUBStack(lStackIndex)
        Else
            Exit Do
        End If

    Loop

    'InsertionSortString sArray, lOldLB, lOldUB
    For i = lOldLB + 1 To lOldUB
        lTemp = m_alStringPtrArray(i)
        
        j = i
        
        Do While j > lOldLB
           m_saStringHeader1.pvData = m_alStringPtrArray(j - 1)
           m_saStringHeader2.pvData = lTemp
           lCharacter = 0
           Do While m_aiString1(lCharacter) And (m_aiString1(lCharacter) = m_aiString2(lCharacter))
              lCharacter = lCharacter + 1
           Loop
           If m_aiString1(lCharacter) = 0 Then lCharacter = lCharacter - 1
           If m_aiString1(lCharacter) <= m_aiString2(lCharacter) Then Exit Do
        
           m_alStringPtrArray(j) = m_alStringPtrArray(j - 1)
           j = j - 1
        Loop
        
        m_alStringPtrArray(j) = lTemp
    Next i
    
    
    'PutMem4 ByVal VarPtr(sTemp), 0
    
End Sub
