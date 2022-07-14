'Main Module -------------------------------------------------------------------
Option Explicit

Sub main()
  UserForm.Show
End Sub


'UserForm -------------------------------------------------------------------

Private targetRng

Private Sub CommandButton_Click()
  ListBox.Clear
  
  Set targetRng = New myRange
  targetRng.setRange RefEdit.Value
  
  Dim field As Range
  
  If CheckBox.Value Then
    For Each field In targetRng.getHeader()
      ListBox.AddItem field.Value
    Next
  Else
    For Each field In targetRng.getHeader()
      ListBox.AddItem Replace(field.EntireColumn.Rows(1).Address, "$1", "列")
    Next
  End If
End Sub

Private Sub CommandButton_Click()
  If ListBox.ListIndex = -1 Then
    TextBox.Value = "キーを選んでください。"
  Else
    TextBox.Value = ListBox.List(ListBox.ListIndex)
    Dim bod As Range
    Debug.Print targetRng.getBody(CheckBox.Value).Address
    
    With ActiveSheet
'      .Sort.SortFields.Clear
'      .Sort.SortFields.Add Key:=.Range("A1"), Order:=xlAscending
 '     .Sort.SortFields.Add Key:=.Range("B1"), Order:=xlAscending
  '    .Sort.setRange mrng.getBody
   '   .Sort.Header = xlYes
    '  .Sort.Apply
    End With
    '.SortFields.Clearメソッドで､前回の情報を消去
    '.SortFields.Addメソッドで､並べ替えのキーを追加
    '.SetRangeで､並べ替えの範囲を設定
    '.Headerで、行ヘッダーの有無を指定
    '.Applyで、並べ替えを実行
  End If
End Sub


'myRange -------------------------------------------------------------------
Option Explicit

Private rng As Range

Public Sub class_initialize()
End Sub

Sub setRange(addr As String)
  Set rng = Range(addr)
End Sub

Function getRange()
  Set getRange = rng
End Function

Function getHeader() As Range
  Set getHeader = rng.Rows(1).Columns()
End Function

Function getBody(includesHeader As Boolean) As Range
  Dim rowLen
  rowLen = rng.Rows.Count
  
  Dim offsetRow
  If includesHeader Then
    offsetRow = 1
  Else
    offsetRow = 0
  End If
  Set getBody = rng.Offset(offsetRow).Resize(rowLen - offsetRow)
  
End Function
