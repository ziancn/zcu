Attribute VB_Name = "ZcuInit"

Sub InitConfigSheet()
    Dim configSheet As Worksheet
    On Error Resume Next
    Set configSheet = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    
    If configSheet Is Nothing Then
        Set configSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        configSheet.Name = "Config"
        
        ' Initialize default configuration values
        configSheet.Range("A1").Value = "Key"
        configSheet.Range("B1").Value = "Value"
        
        ' Add default configuration entries
        configSheet.Range("A2").Value = "ExampleKey"
        configSheet.Range("B2").Value = "ExampleValue"
    End If
End Sub