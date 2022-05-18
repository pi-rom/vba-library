Attribute VB_Name = "stringUtilities"
Function charSplit(ByVal str As String) As String()
    Dim output As Variant
    str = StrConv(str, vbUnicode)
    output = Split(str, vbNullChar)
    ReDim Preserve output(UBound(output) - 1)
    charSplit = output
End Function

