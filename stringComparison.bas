Attribute VB_Name = "stringComparison"
Function hammingDistance(ByVal str1 As String, ByVal str2 As String) As Integer

    Dim i, count As Integer
    Dim strArray1, strArray2 As Variant
    i = 0
    count = 0

    strArray1 = charSplit(str1)
    strArray2 = charSplit(str2)

    For i = 0 To UBound(strArray1)
        If strArray1(i) <> strArray2(i) Then count = count + 1
    Next i

    hammingDistance = count

End Function

Private Sub levMatrix( _
        ByRef strArray1 As Variant, _
        ByRef strArray2 As Variant, _
        ByRef str1 As String, _
        ByRef str2 As String, _
        ByRef d As Variant _
        )

    strArray1 = charSplit(str1)
    strArray2 = charSplit(str2)
    
    ReDim d(UBound(strArray1) + 1, UBound(strArray2) + 1)
    
    For i = 0 To UBound(strArray1) + 1
        d(i, 0) = i
    Next
    
    For j = 0 To UBound(strArray2) + 1
        d(0, j) = j
    Next

End Sub

Function levDistance(ByVal str1 As String, ByVal str2 As String) As Integer

    Dim i, j, min1, min2, min3, cost, d() As Integer
    Dim strArray1, strArray2 As Variant

    Call levMatrix(strArray1, strArray2, str1, str2, d)
    
    If UBound(strArray1) = 0 Then
        levDistance = UBound(strArray2)
        Exit Function
    End If
    
    If UBound(strArray2) = 0 Then
        levDistance = UBound(strArray1)
        Exit Function
    End If
    
    For i = 1 To UBound(strArray1) + 1
        For j = 1 To UBound(strArray2) + 1
            If strArray1(i - 1) = strArray2(j - 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            min1 = (d(i - 1, j) + 1)
            min2 = (d(i, j - 1) + 1)
            min3 = (d(i - 1, j - 1) + cost)
            
            d(i, j) = Application.Min(min1, min2, min3)
            
        Next j
    Next i
    
    levDistance = d(UBound(strArray1) + 1, UBound(strArray2) + 1)
End Function

Function damLevDistance(ByVal str1 As String, ByVal str2 As String) As Integer

    Dim i, j, min1, min2, min3, cost, d() As Integer
    Dim strArray1, strArray2 As Variant

    Call levMatrix(strArray1, strArray2, str1, str2, d)
    
    If UBound(strArray1) = 0 Then
        damLevDistance = UBound(strArray2)
        Exit Function
    End If
    
    If UBound(strArray2) = 0 Then
        damLevDistance = UBound(strArray1)
        Exit Function
    End If
    
    For i = 1 To UBound(strArray1) + 1
        For j = 1 To UBound(strArray2) + 1
            If strArray1(i - 1) = strArray2(j - 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            min1 = (d(i - 1, j) + 1)
            min2 = (d(i, j - 1) + 1)
            min3 = (d(i - 1, j - 1) + cost)
            
            d(i, j) = Application.Min(min1, min2, min3)
            
            '-------------------
            'Damerau's algorithm
            '-------------------
            If i > 1 And j > 1 Then
                If strArray1(i - 2) = strArray2(j - 1) And strArray1(i - 1) = strArray2(j - 2) Then
                    d(i, j) = Application.Min(d(i, j), d(i - 2, j - 2) + cost)
                End If
            End If
        Next j
    Next i
    
    damLevDistance = d(UBound(strArray1) + 1, UBound(strArray2) + 1)
End Function


