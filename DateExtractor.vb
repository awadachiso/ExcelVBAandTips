Option Explicit

Private re As RegExp
Private patterns As Collection
Private dateOfExtraction 'result

Public Sub class_initialize()
  Set re = New RegExp
  re.Global = True
  
  Set patterns = New Collection
  setPatterns
End Sub

Public Sub class_terminate()
  Set re = Nothing
End Sub

Public Sub setPatterns()
  'target is "yyyyXmmXdd"
  ' X may contain symbols such as "/" , "月" , "日",  or " "
  'and they are unnecessary.
  patterns.Add "(20\d{2})\/?年? ?0?([1-9]|1[0-2])\/?月? ?0?([1-2]?[1-9]|3[0-1])日?"
  
  'When corresponding to a new type Date string, add new  pattern to collect in order of (year), (month), and (date).
End Sub


Public Function matchAnyPattern(target As String)
  Dim result As Boolean
  result = False
  
  Dim pat As Variant
  For Each pat In patterns
    re.Pattern = pat
    result = result Or re.test(target)
  Next
  matchAnyPattern = result
End Function

Private Sub convertToCorrectDateExpression(target As String)
  Dim pat As Variant
  Dim matches As MatchCollection
  Dim result
  result = ""
    
  For Each pat In patterns
    re.Pattern = pat
    If re.test(target) Then
        Set matches = re.Execute(target)
        
        Dim sMatch
        'Ignore all matches but the first.
        'Loop ()matches in regexp
        For Each sMatch In matches(0).SubMatches
          If (result = "") Then
            result = sMatch
          Else
            result = result & "/" & sMatch
          End If
        Next
        
      End If
  Next
  
  dateOfExtraction = result
End Sub


Public Function getDate(target)
  convertToCorrectDateExpression (target)
  'return date as value
  getDate = DateValue(dateOfExtraction)
End Function
