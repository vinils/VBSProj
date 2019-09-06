Sub ArrayAppend(ByRef arr, ByVal arr0)
    Dim i, ub
    If IsArray(arr) And IsArray(arr0) Then
        On Error Resume Next
        ub = UBound(arr)
        If Err.Number <> 0 Then ub = -1
        ReDim Preserve arr(ub + UBound(arr0))
        For i = 0 To UBound(arr0)
            arr(ub + 1 + i) = arr0(i)
        Next
    End If
End Sub
