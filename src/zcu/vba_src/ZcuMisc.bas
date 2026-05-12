Attribute VB_Name = "ZcuMisc"

Public Function GetCurrentFileDirectory() As String
    GetCurrentFileDirectory = ThisWorkbook.Path
End Function

Public Function GetDefaultUvPythonPath() As String
    currentDir = ThisWorkbook.Path
    GetDefaultUvPythonPath = currentDir & "\.venv\Scripts\python.exe"
End Function

