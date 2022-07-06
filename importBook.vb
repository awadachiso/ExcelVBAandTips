Option Explicit

Function importBook(srcBook As Workbook, destBook As Workbook, Optional sheetNum As Integer = 0)
  
  If TypeName(srcBook) = "Workbook" Then
    Debug.Print "<- " & srcBook.Name
  Else
    Debug.Print "<- nothing"
  End If
  
  If TypeName(destBook) = "Workbook" Then
    Debug.Print "-> " & destBook.Name
  Else
    Debug.Print "-> nothing"
  End If
  
  
End Function

Sub test()
  Dim srcFileNames As Variant, srcName
  Dim srcBook As Workbook
  Dim result
  
  Call disableAutoOperationBeforeExecMacro
  
  srcFileNames = getFileNamesFromDialog()
  
  On Error GoTo FailToOpenFile
  ActiveWindow.Visible = False
  For Each srcName In srcFileNames
    Set srcBook = Workbooks.Open(srcName)
    Debug.Print srcBook.Name
    srcBook.Close False
    'result = importBook(srcBook, ThisWorkbook)
  Next
  Exit Sub
    
  Call enableAutoOperationAfterExecMacro
FailToOpenFile:
  MsgBox Prompt:=Err.Description, Title:="Error Occured"
  Call enableAutoOperationAfterExecMacro
End Sub
