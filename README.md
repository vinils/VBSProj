# VBSProj
VBScripts project with Test Cases, CI and quality gate

[![Build Status](https://travis-ci.org/vinils/VBSProj.svg?branch=master)](https://travis-ci.org/vinils/VBSProj)

# Samples
- StrFormart("My string {0} format", Array("Test"))
- ArrayApend(Array1ByRef, ArrByVal)

# Hot to use it
```vbscript
Sub Import(strFile)
    strFile = strFile + ".vbs"

    Dim objFile, strCode, objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strFile)
    strCode = objFile.ReadAll
    objFile.Close
    ExecuteGlobal strCode
    Set objFile = Nothing
    Set objFSO = Nothing
End Sub

Import "StrFormat"
Import "ArrayApend"

MsgBox(StrFormart("My string {0} format", Array("Test")))
```
