Option Explicit

Function getFileNamesFromDialog()
  Dim myFiles As Variant
  Dim f As Variant
  
  ChDir "C:\Users" 'select default dir

  myFiles = Application.GetOpenFilename( _
    FileFilter:="Excel (*.xlsx;*.xls;*.xls),*.xlsx;*.xls;*.xls, CSV (*.csv),*.csv", _
    MultiSelect:=True)
  
  'Check if the file was not selected
  If isEmpty(myFiles) Then
    getFileNamesFromDialog = myFiles
  Else
    'IF Not Selected, returns empty Array
    getFileNamesFromDialog = Array()
  End If
End Function


Function isEmpty(arr)
  isEmpty = IsArray(arr)
End Function


Sub test()
  Dim files As Variant
  Dim f As Variant
  files = getFileNamesFromDialog()
  
  For Each f In files
    Debug.Print f
  Next
End Sub
