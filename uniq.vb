Option Explicit

Function uniq(rngA, rngB)
  Dim result
  Set result = CreateObject("Scripting.Dictionary")
    
  Dim integratedRange, length, cll
  Set integratedRange = Union(rngA, rngB)
  length = integratedRange.Count
  
  'add first cell
  result.Add integratedRange(1).Value, integratedRange(1).Value
  
  'if not exist, add to result
  'Key and Item are cell.value
  For Each cll In integratedRange
    If Not result.exists(cll.Value) Then
      result.Add cll.Value, cll.Value
    End If
  Next
  
  'length is sizeof integratedRange
  'In order to be used as an array expression, it must be at least as large as integratedRange.
  'Therefore, the keys of the remaining elements will be the current size,
  'and the values will be padded with "".
  Dim i
  For i = result.Count To length
    result.Add result.Count, ""
  Next
  
  uniq = WorksheetFunction.Transpose(result.items)
  result = Null
End Function
