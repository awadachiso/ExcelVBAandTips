Option Explicit

Private filePathList As Collection
Private index

Public Sub class_initialize()
  Set filePathList = New Collection
  index = 1 'Collection is 1-based.
End Sub

Public Sub addFiles()
  Dim FD As FileDialog 'need to set refference Microsoft Scripting Runtime'
  Set FD = Application.FileDialog(msoFileDialogOpen)
  With FD
      'set the target extentions
      .Filters.add "src files", "*.xls; *.xlsx; *.xlsm; *.csv", 1
      'change default directory
      .InitialFileName = "C:\Users\" & Environ("USERNAME") & "\Downloads\"
      
      If .Show = True Then ' if valid operation
        'register Selected files to filePathList
        Dim itm
        For Each itm In .SelectedItems
          filePathList.add itm
        Next
      Else ' if INvalid operation
          'cancel button pressed
      End If
    End With
End Sub


Public Function nextItem()
  nextItem = filePathList(index)
  index = index + 1
End Function

Public Function hasNext() As Boolean
  hasNext = index <= filePathList.Count
End Function
