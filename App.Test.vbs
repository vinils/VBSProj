Dim oFSO, oFile, strCode

Const ForReading = 1
Const ForWriting = 2

Set oFSO = CreateObject("Scripting.FileSystemObject")
sFolder = oFSO.GetAbsolutePathName(".")

For Each oFile In oFSO.GetFolder(sFolder).Files
  If UCase(oFile.Name) <> "APP.TEST.VBS" Then
    If UCase(Right(oFile.Name, 8)) = "TEST.VBS" Then
      WSCript.Echo "Executing " & oFile.Name
      Set oFile = oFSO.OpenTextFile(oFile.Name)
      strCode = oFile.ReadAll
      oFile.Close
      Execute strCode
      Set objFile = Nothing
      Set objFSO = Nothing
      Set oFSO = Nothing
    End if
  End if
Next

Set oFSO = Nothing
