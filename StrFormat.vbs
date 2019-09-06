Public Function StrFormat(FormatString, Arguments())
    Dim CurArgNum

    StrFormat = FormatString

    For CurArgNum = UBound(Arguments) To 0 Step -1
        StrFormat = Replace(StrFormat, "{" & CurArgNum & "}", Arguments(CurArgNum))
    Next
End Function