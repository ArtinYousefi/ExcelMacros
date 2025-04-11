Function WordCount(ParamArray cells() As Variant) As String
    Dim result As String
    Dim i As Long
    Dim txt As String
    Dim wordArray() As String
    Dim count As Long

    result = ""

    For i = LBound(cells) To UBound(cells)
        If TypeName(cells(i)) = "Range" Then
            txt = Trim(cells(i).Value)
            If Len(txt) = 0 Then
                count = 0
            Else
                wordArray = Split(txt)
                count = UBound(wordArray) - LBound(wordArray) + 1
            End If
            result = result & count & ", "
        End If
    Next i

    ' Remove trailing comma and space
    If Len(result) > 2 Then
        result = Left(result, Len(result) - 2)
    End If

    WordCount = result
End Function
