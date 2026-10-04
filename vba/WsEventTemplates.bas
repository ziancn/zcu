' Worksheet Event Templates

''' LIFECYCLE
Private Sub Worksheet_Activate()
End Sub

Private Sub Worksheet_Deactivate()
End Sub

Private Sub Worksheet_BeforeDelete()
End Sub


''' CELL / USER INTERACTION
Private Sub Worksheet_Change(ByVal Target As Range)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
End Sub


''' CALCULATION
Private Sub Worksheet_Calculate()
End Sub


''' HYPERLINKS
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
End Sub


''' PIVOT TABLES
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
End Sub

Private Sub Worksheet_PivotTableAfterValueChange(ByVal TargetPivotTable As PivotTable, ByVal TargetRange As Range)
End Sub

Private Sub Worksheet_PivotTableBeforeAllocateChanges( _
    ByVal TargetPivotTable As PivotTable, _
    ByVal ValueChangeStart As Long, _
    ByVal ValueChangeEnd As Long, _
    Cancel As Boolean)
End Sub

Private Sub Worksheet_PivotTableBeforeCommitChanges( _
    ByVal TargetPivotTable As PivotTable, _
    ByVal ValueChangeStart As Long, _
    ByVal ValueChangeEnd As Long, _
    Cancel As Boolean)
End Sub

Private Sub Worksheet_PivotTableBeforeDiscardChanges( _
    ByVal TargetPivotTable As PivotTable, _
    ByVal ValueChangeStart As Long, _
    ByVal ValueChangeEnd As Long)
End Sub

Private Sub Worksheet_PivotTableChangeSync(ByVal Target As PivotTable)
End Sub