'''
' === Key Worksheet Events ===
' Worksheet_Change
' Worksheet_SelectionChange
' Worksheet_BeforeRightClick
' Worksheet_BeforeDoubleClick
' Worksheet_Activate
' Worksheet_Deactivate
' Worksheet_Calculate
' Worksheet_FollowHyperlink
' Worksheet_PivotTableUpdate
' === Key Workbook Events ===
' 
' === Key Application Events ===
' 
'''


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim tbl As ListObject
    Dim rowData As Range
   
    ' --- DETECT TABLE OBJECT---
    On Error Resume Next
    Set tbl = Target.ListObject
    On Error GoTo 0
   
    ' If not in a table, do nothing
    If tbl Is Nothing Then Exit Sub
   
    ' --- FILTER BY SPECIFIC COLUMN ---
    ' ### Change 2 to other index or column name you want ###
    If Intersect(Target, tbl.ListColumns(2).DataBodyRange) Is Nothing Then Exit Sub
   
    ' --- PROCESS THE DATA ---
    If Not Intersect(Target, tbl.DataBodyRange) Is Nothing Then
        Set rowData = Intersect(Target.EntireRow, tbl.Range)
       
        ' Cancel the default cell-editing behavior
        Cancel = True
       
        ' Call YourActionProcedure(rowData, tbl.HeaderRowRange)
        MsgBox "Double-clicked " & Target.Value & " in column 2 of " & tbl.Name
    End If
End Sub



