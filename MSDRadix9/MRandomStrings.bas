Attribute VB_Name = "MRandomStrings"
'***********************************************************************************************
'
'    Module:            MRandomStrings [MRandomStrings.bas]
'    Author:            Ralph Eastwood
'    Email:             tcmreastwood@ntlworld.com
'    Creation:          Sunday 24, December 2006, 7:29:45 AM
'    Last Modification: Sunday 24, December 2006, 7:29:45 AM
'    Purpose:           Generate Random Strings
'    Notes:
'    Update History:    - Created (24/12/2006)
'
'***********************************************************************************************

Option Explicit


Public Sub GenerateRandomStrings(ByVal lLength As Long, ByRef lCount As Long, ByRef asOutString() As String, Optional ByVal lRandomStringLengthDeviation As Long = 0)

    Dim i As Long, j As Long
    Dim lCalculatedLength  As Long
    Dim strChar As String
    
    ReDim asOutString(lCount - 1)
    
    For i = 0 To lCount - 1
        lCalculatedLength = lLength + Int((lRandomStringLengthDeviation + lRandomStringLengthDeviation + 1) * Rnd - lRandomStringLengthDeviation)
        asOutString(i) = Space$(lCalculatedLength)
        For j = 1 To lCalculatedLength
            strChar = Chr$(Int(Rnd * 50) + 60)
            Mid$(asOutString(i), j, 1) = strChar
        Next j
    Next i

End Sub

Public Sub RandomiseList(ByRef asOutString() As String)

    Dim lLBound As Long, lUBound As Long
    Dim i As Long, r As Long, sTmp As String
    
    lLBound = LBound(asOutString)
    lUBound = UBound(asOutString)

    For i = lLBound To lUBound
        r = Int((lUBound - lLBound + 1) * Rnd + lLBound)
        sTmp = asOutString(i)
        asOutString(i) = asOutString(r)
        asOutString(r) = sTmp
    Next i

End Sub

Public Function CheckSort(ByRef asList() As String, ByVal lCount As Long, Optional ByVal fDescending As Boolean = False) As Long
    
    Dim i As Long
    
    For i = 0 To lCount - 2
        If fDescending Then
            If asList(i) < asList(i + 1) Then
                CheckSort = i
                Exit Function
            End If
        Else
            If asList(i) > asList(i + 1) Then
                CheckSort = i
                Exit Function
            End If
        End If
    Next i
    CheckSort = -1
    
End Function
