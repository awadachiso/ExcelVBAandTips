Option Explicit

Sub main()
  Dim i, j As Integer
  For i = 1 To Selection.Count - 1
    j = Selection.Count - i + 1
    Debug.Print Selection(j).Address
    If Selection(j).Value <> Selection(j - 1).Value Then
      Selection(j).EntireRow.Insert
    End If
  Next
End Sub


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

Sub rollup(col)
  Dim length
  length = col.Rows.Count
  Dim i
  i = length
  
  While i > 0
    If Not isSameValueWithUpperCell(col(i)) Then
      If i = 1 Then
 '       insertUpperRow col(i)
      ElseIf col(i).Offset(-1).Value <> "" Then
        insertUpperRow col(i)
      End If
    Else
      col(i).ClearContents
    End If
    i = i - 1
  Wend
End Sub

Sub foldleft(rng)
  Dim width
  width = rng.Columns.Count
  Dim col
  For Each col In rng.Columns
    If isLastColumn(col, rng) Then
      'do nothing
    Else
      Call rollup(col)
    End If
  Next
End Sub

Function isLastColumn(col, rng)
  Dim result
  result = col.Column = rng.Columns(rng.Columns.Count).Column
  isLastColumn = result
End Function