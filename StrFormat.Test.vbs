Sub Import(strFile)
    strFile = strFile + ".vbs"
    strFile = Replace(strFile, ".\", $GetAppPath())

    Dim objFile, strCode, objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strFile)
    strCode = objFile.ReadAll
    objFile.Close
    ExecuteGlobal strCode
    Set objFile = Nothing
    Set objFSO = Nothing
End Sub

Import ".\Tests"
Import ".\StrFormat"

Function TC01_StringFormat_WhenHasParameters_ShouldBeReplaced()
    Dim stringFormat, input, result
    
    'Arrange
    stringFormat = "My strFormat with {0} content"
    input = "number 1"

    'Act
    result = StrFormat(stringFormat, Array(input))

    'Assert
    Assert result = "My strFormat with number 1 content", "Strformat My strFormat with number {0} content not replaced properly"
End Function

WScript.Echo "Testing - TC01_StringFormatHasWithInput_WhenStrFormat_ShouldBeReplaced"
TC01_StringFormatHasWithInput_WhenStrFormat_ShouldBeReplaced()
