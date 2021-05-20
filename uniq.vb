Option Explicit

Function uniq(rngA, rngB)
  Dim uniqs
  Set uniqs = CreateObject("Scripting.Dictionary")
    
  Dim maxLength, cll
  maxLength = rngA.Count + rngB.Count
  
  'add first cell
  uniqs.Add rngA(1).Value, rngA(1).Value
  
  'if not exist, add to uniqs
  'Key and Item are cell.value
  For Each cll In rngA
    If cll.Value <> "" Then
      If Not uniqs.exists(cll.Value) Then
        uniqs.Add cll.Value, ""
      End If
    End If
  Next
  
  For Each cll In rngB
    If cll.Value <> "" Then
      If Not uniqs.exists(cll.Value) Then
        uniqs.Add cll.Value, ""
      End If
    End If
  Next
  
  'maxLength is sizeof rngA + rngB
  'In order to be used as an array expression, size of retuern value must be at least maxLength.
  'Therefore, the remaining elements will be padded with "".
  
  'extract Keys
  Dim result
  result = uniqs.Keys
  'resize
  ReDim Preserve result(maxLength)
  
  Dim i
  For i = uniqs.Count To maxLength
    result(i) = ""
  Next
  
  uniq = WorksheetFunction.Transpose(result)
  uniqs = Null
End Function
