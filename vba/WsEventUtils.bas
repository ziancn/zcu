Option Explicit

Public Sub ApplyTableHeaderFilter(ByVal Target As Range)
    ' Apply single-cell target text as a filter to the corresponding table column header
    Dim headerCell As Range
    Dim lo As ListObject
    Dim fieldIndex As Long
    Dim criteria As String
    Dim criteria1 As String
    Dim criteria2 As String
    Dim operator As XlAutoFilterOperator

    If Target.Cells.CountLarge <> 1 Then Exit Sub

    Set headerCell = Target.Offset(1, 0)
    If headerCell.ListObject Is Nothing Then Exit Sub

    Set lo = headerCell.ListObject
    If Intersect(headerCell, lo.HeaderRowRange) Is Nothing Then Exit Sub

    fieldIndex = headerCell.Column - lo.Range.Column + 1
    criteria = Trim$(CStr(Target.Value2))

    If Len(criteria) = 0 Then
        lo.Range.AutoFilter Field:=fieldIndex   ' Clear filter when input cell is empty
        Exit Sub
    End If

    If Not ParseFilterCriteria(criteria, criteria1, criteria2, operator) Then Exit Sub

    If Len(criteria2) > 0 Then
        lo.Range.AutoFilter Field:=fieldIndex, Criteria1:=criteria1, Operator:=operator, Criteria2:=criteria2
    Else
        lo.Range.AutoFilter Field:=fieldIndex, Criteria1:=criteria1
    End If

End Sub


Private Function ParseFilterCriteria( _
    ByVal inputText As String, _
    ByRef criteria1 As String, _
    ByRef criteria2 As String, _
    ByRef operator As XlAutoFilterOperator) As Boolean
    ' Parse the input text to determine the appropriate filter criteria and operator for AutoFilter.
    Dim leftBracket As String
    Dim rightBracket As String
    Dim commaPos As Long
    Dim lowerBound As String
    Dim upperBound As String

    inputText = Trim$(inputText)

    criteria1 = vbNullString
    criteria2 = vbNullString
    operator = 0

    ' Exact text: "ABC"
    If Len(inputText) >= 2 Then
        If Left$(inputText, 1) = """" And Right$(inputText, 1) = """" Then
            criteria1 = Mid$(inputText, 2, Len(inputText) - 2)
            ParseFilterCriteria = True
            Exit Function
        End If
    End If

    ' Range: [25,100), (25,100], [25,100], (25,100)
    If (Left$(inputText, 1) = "[" Or Left$(inputText, 1) = "(") And _
       (Right$(inputText, 1) = "]" Or Right$(inputText, 1) = ")") Then

        leftBracket = Left$(inputText, 1)
        rightBracket = Right$(inputText, 1)
        commaPos = InStr(2, inputText, ",")

        If commaPos = 0 Then Exit Function

        lowerBound = Trim$(Mid$(inputText, 2, commaPos - 2))
        upperBound = Trim$(Mid$(inputText, commaPos + 1, Len(inputText) - commaPos - 1))

        If Not IsNumeric(lowerBound) Or Not IsNumeric(upperBound) Then Exit Function

        If leftBracket = "[" Then
            criteria1 = ">=" & lowerBound
        Else
            criteria1 = ">" & lowerBound
        End If

        If rightBracket = "]" Then
            criteria2 = "<=" & upperBound
        Else
            criteria2 = "<" & upperBound
        End If

        operator = xlAnd
        ParseFilterCriteria = True
        Exit Function

    End If

    ' >=, <=, <>
    If Left$(inputText, 2) = ">=" Or _
       Left$(inputText, 2) = "<=" Or _
       Left$(inputText, 2) = "<>" Then

        criteria1 = Left$(inputText, 2) & Trim$(Mid$(inputText, 3))
        ParseFilterCriteria = True
        Exit Function

    End If

    ' >, <, =
    If Left$(inputText, 1) = ">" Or _
       Left$(inputText, 1) = "<" Or _
       Left$(inputText, 1) = "=" Then

        criteria1 = Left$(inputText, 1) & Trim$(Mid$(inputText, 2))
        ParseFilterCriteria = True
        Exit Function

    End If

    ' Numeric value: exact match
    If IsNumeric(inputText) Then
        criteria1 = inputText
        ParseFilterCriteria = True
        Exit Function
    End If

    ' Text: contains by default
    If InStr(inputText, "*") = 0 And InStr(inputText, "?") = 0 Then
        criteria1 = "*" & inputText & "*"
    Else
        criteria1 = inputText
    End If

    ParseFilterCriteria = True

End Function