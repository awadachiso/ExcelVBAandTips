Function uniq(rngA, rngB)
  Dim result
  Set result = CreateObject("Scripting.Dictionary")
  
  result.Add rngA(1).Value, rngA(1)
  
  Dim a, b
  For Each a In rngA
    If Not result.exists(a.Value) Then
      result.Add a.Value, a
      Debug.Print a.Value, a.Address
    End If
  Next
  
  For Each b In rngB
   If Not result.exists(b.Value) Then
     result.Add b.Value, b
     Debug.Print b.Value, b.Address
   End If
  Next
  
  uniq = WorksheetFunction.Transpose(result.items)
  result = Null
End Function