Option Explicit

'switch src Sheets By blockName
'e.g.) =sumBlock("A1:A10", "BlockA") using Spill or ArrayFromula(CSE) 
Function sumBlock(rngStr As String, blockName As String)
  Dim shts
  Dim sheetNames
  Set sheetNames = CreateObject("Scripting.Dictionary")
  sheetNames.Add "BlockA", Array(ThisWorkbook.Sheets("Sheet1"), ThisWorkbook.Sheets("Sheet2"))
  sheetNames.Add "BlockB", Array(ThisWorkbook.Sheets("Sheet3"), ThisWorkbook.Sheets("Sheet4"))
  
  shts = sheetNames(blockName)
  sumBlock = sumBetweenSheets(shts, Range(rngStr))
End Function


Function averageBlock(rngStr As String, blockName As String)
  Dim shts
  Dim sheetNames
  Set sheetNames = CreateObject("Scripting.Dictionary")
  sheetNames.Add "BlockA", Array(ThisWorkbook.Sheets("Sheet1"), ThisWorkbook.Sheets("Sheet2"))
  sheetNames.Add "BlockB", Array(ThisWorkbook.Sheets("Sheet3"), ThisWorkbook.Sheets("Sheet4"))
  
  shts = sheetNames(blockName)
  averageBlock = averageBetweenSheets(shts, Range(rngStr))
End Function

Function sumBetweenSheets(shts, rng As Range)
  Dim results() As Variant
  Dim height: height = rng.Rows.Count
  Dim width: width = rng.Columns.Count
  
  'Make the size of the results() the same as that of the argument
  ReDim results(height - 1, width - 1)
  
  'Find the total value of all sheets in each cell
  Dim x, y  'counter
  Dim s     'sheet iterator
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
    Next
  Next
  
  'return Array
  sumBetweenSheets = results
End Function


Function averageBetweenSheets(shts, rng As Range)
  Dim results() As Variant
  Dim height: height = rng.Rows.Count
  Dim width: width = rng.Columns.Count
  
  'Make the size of the results() the same as that of the argument
  ReDim results(height - 1, width - 1)
  
  'Find the total value of all sheets in each cell
  Dim x, y  'counter
  Dim s     'sheet iterator
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
       results(y, x) = results(y, x) / Sheets.Count
    Next
  Next
  
  'return Array
  averageBetweenSheets = results
End Function

