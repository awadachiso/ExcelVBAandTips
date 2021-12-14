
Function isSameValueWithUpperCell(cll)
  Dim result As Boolean
  result = False
  
  If cll.Row() = 1 Then
    result = False
  Else
    Dim upperCell
    Set upperCell = cll.Offset(-1)
    result = cll.Value = upperCell.Value
  End If
    
  isSameValueWithUpperCell = result
End Function

Sub insertUpperRow(cll)
  cll.EntireRow.Insert
  cll.Offset(-1).Value = cll.Value
End Sub

Sub rollup(Col)
  Dim length
  length = Col.Rows.Count
  Dim rowNum
  rowNum = length
  Dim currentCell
  
  
  While rowNum > 0
    Set currentCell = Col(rowNum)
    If currentCell.Value = "" Then
      'Do Nothing
    ElseIf rowNum <> 1 And isSameValueWithUpperCell(currentCell) Then
      'Do Nothing
      'currentCell.ClearContents
    Else
      insertUpperRow currentCell
      'currentCell.ClearContents
    End If
    rowNum = rowNum - 1
  Wend
End Sub

Sub foldleft(rng)
  Dim width
  width = rng.Columns.Count
  Dim Col
  For Each Col In rng.Columns
    If isLastColumn(Col, rng) Then
      'do nothing
    Else
      Call rollup(Col)
    End If
  Next
End Sub

Function isLastColumn(Col, rng)
  Dim result
  result = Col.Column = rng.Columns(rng.Columns.Count).Column
  isLastColumn = result
End Function

Sub m2(Col)
  rollup Range(Col & "1:" & Col & "100 ")
End Sub
