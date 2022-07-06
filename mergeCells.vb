Option Explicit

Sub mergeSequentialCellsWithSameValue()
  Application.DisplayAlerts = False
  
  Dim targetRng As Range, initialCell As Range, mergingRng As Range
  Dim rngType As String ' "vertical" or "hoirzontal"
  Dim currentVal
  Dim length As Integer
  
  
  If Selection.Count = 1 Then ' if select single cell
    Set targetRng = ActiveSheet.UsedRange
  Else
    Set targetRng = Selection
  End If
  
  length = targetRng.Count
  
  On Error GoTo errorHandle
    
  'check the range shape
  If targetRng.Rows.Count = 1 Then
    rngType = "horizontal"
  ElseIf targetRng.Columns.Count = 1 Then
    rngType = "vertical"
  Else ' if Not single row or column
      Error.Raise
  End If
    
  Set initialCell = targetRng(1)
  currentVal = initialCell.Value
  
  Dim r
  For Each r In targetRng
    If r.Value <> currentVal Then ' if found different value
      'merge initialCell with the previous cell
      Call mergeInitWithPreviousCell(initialCell, r, rngType)
      
      'resetting initialCell and value
      Set initialCell = r
      currentVal = r.Value
    End If
  Next
  Exit Sub
  
errorHandle:
  MsgBox "select single row or column"
  
End Sub


Sub mergeSequentialEmptyCells()
  Application.DisplayAlerts = False
  
  '対象範囲, 起点セル,結合範囲
  Dim targetRng As Range, initialCell As Range
  Dim rngType As String ' "vertical" or "hoirzontal"
  Dim currentVal
  Dim length As Integer
  
  
  If Selection.Count = 1 Then ' if select single cell
    Set targetRng = ActiveSheet.UsedRange
  Else
    Set targetRng = Selection
  End If
  
  length = targetRng.Count
  
  On Error GoTo errorHandle
    
  'check the range shape
  If targetRng.Rows.Count = 1 Then
    rngType = "horizontal"
  ElseIf targetRng.Columns.Count = 1 Then
    rngType = "vertical"
  Else ' if Not single row or column
      Err.Raise 999, Description:="select single row or column"
  End If
    
  Set initialCell = targetRng(1)
  currentVal = initialCell.Value
  
  If currentVal = "" Then
    Err.Raise 991, Description:="initialCell is empty"
  End If
  
  Dim r
  For Each r In targetRng
    If r.Value <> currentVal And r.Value <> "" Then  ' if found different value
'     Call mergeInitWithPreviousCell(initialCell, r, rngType)
'    'resetting initialCell and value
'     Set initialCell = r
'     currentVal = r.Value
    ElseIf isLastCell(targetRng, r) Then  ' if found end of range
      Call mergeInitWithCurrentCell(initialCell, r)
     'resetting initialCell and value
      Set initialCell = r
      currentVal = r.Value
    End If
  Next
  Exit Sub
  
errorHandle:
  MsgBox Err.Description
  
End Sub


Sub mergeInitWithCurrentCell(initialCell, currentCell)
    Range(initialCell, currentCell).Merge
End Sub

Sub mergeInitWithPreviousCell(initialCell, currentCell, rngType)
    If rngType = "vertical" Then
      Call mergeInitWithCurrentCell(initialCell, currentCell.Offset(-1, 0))
    ElseIf rngType = "horizontal" Then
      Call mergeInitWithCurrentCell(initialCell, currentCell.Offset(0, -1))
    End If
End Sub

Function isLastCell(rng, cll)
  isLastCell = rng(rng.Count) = cll
End Function
