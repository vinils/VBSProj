Set Console = CreateObject("Scripting.FileSystemObject")
Console.GetStandardStream(1).WriteLine "https://helloacm.com"

WSCript.Echo StrFormat("my test {0}", Array("asdf"))

Assert 20 < 15, "pl1 < pl2"

Sub Assert(x, msg)
    If Not x Then
        ' alternatively, you can throw Errors using Err.Raise
        'Console.GetStandardStream(1).WriteLine msg 
        Err.Raise 1, msg, msg
    End If
End Sub

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

Public Function StrFormat(FormatString, Arguments())
    Dim CurArgNum

    StrFormat = FormatString

    For CurArgNum = UBound(Arguments) To 0 Step -1
        StrFormat = Replace(StrFormat, "{" & CurArgNum & "}", Arguments(CurArgNum))
    Next
End Function
