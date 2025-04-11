Function CountWordsOverVerbose(cutoff As Long, ParamArray cells() As Variant) As String
    Dim i As Long
    Dim txt As String
    Dim wordArray() As String
    Dim count As Long
    Dim totalOver As Long
    Dim overList As String

    totalOver = 0
    overList = ""

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
                overList = overList & count & ", "
            End If
        End If
    Next i

    If Len(overList) > 2 Then
        overList = Left(overList, Len(overList) - 2) ' remove trailing comma + space
    End If

    CountWordsOverVerbose = "Count: " & totalOver & " — Over " & cutoff & ": " & overList
End Function

