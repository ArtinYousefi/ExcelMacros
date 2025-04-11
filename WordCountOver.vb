Function CountWordsOver(cutoff As Long, ParamArray cells() As Variant) As Long
    Dim i As Long
    Dim txt As String
    Dim wordArray() As String
    Dim count As Long
    Dim totalOver As Long

    totalOver = 0

    For i = LBound(cells) To UBound(cells)
        If TypeName(cells(i)) = "Range" Then
            txt = Trim(cells(i).Value)
            If Len(txt) = 0 Then
                count = 0
            Else
                wordArray = Split(txt)
                count = UBound(wordArray) - LBound(wordArray) + 1
            End If
            If count > cutoff Then
                totalOver = totalOver + 1
            End If
        End If
    Next i

    CountWordsOver = totalOver
End Function

