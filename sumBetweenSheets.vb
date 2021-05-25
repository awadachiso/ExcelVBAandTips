Option Explicit

Function sumOneTwo(rng As Range)
  Dim shts
  shts = Array(ThisWorkbook.Sheets("Sheet1"), ThisWorkbook.Sheets("Sheet2"))
  sumOneTwo = sumMultiSheet(shts, rng)
End Function

Function SumValueBetweenSheets(shts, rng As Range)
  Dim results() As Variant
  Dim height: height = rng.Rows.Count
  Dim width: width = rng.Columns.Count
  
  'Make the size of the results() the same as that of the argument
  ReDim results(height - 1, width - 1)
  
  'Find the total value of all sheets in each cell
  Dim x, y  'counter
  dim s     'sheet iterator
  Dim cellIndex
  
  'Tally while proceeding in the direction of the column
  For y = 0 To height - 1 'Vertical (row direction)
    For x = 0 To width - 1 'Horizontal (column direction)
      'rowNum*width +colNum+1
      cellIndex = y * width + x + 1
      'initialize
      results(y, x) = 0
      
      For Each s In shts
        results(y, x) = results(y, x) + s.Range(rng(cellIndex).Address)
      Next
      
      ' case of average
      ' For Each s In Sheets
        ' results(y, x) = results(y, x) + s.Range(rng(cellIndex).Address)
      ' Next
      ' results(y, x) = results(y, x) / Sheets.Count
      
    Next
  Next
  
  'return Array
  sumMultiSheet = results
End Function

Function createSheetsArray(ParamArray sheetNames() As Variant)
  Dim collect
  Set collect = New Collection

  Dim results
  Dim maxIndex
  maxIndex = UBound(sheetNames)
  ReDim results(maxIndex)
    
  Dim i
  For i = 0 To maxIndex
    collect.Add ThisWorkbook.Sheets(sheetNames(i)), sheetNames(i)
  Next
  
  For Each c In collect
    
  
  createSheetsArray = collect
End Function
