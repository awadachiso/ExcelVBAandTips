Option Explicit

Private filePathList As Collection
Private index

Public Sub class_initialize()
  Set filePathList = New Collection
  index = 1
End Sub

Public Sub addFiles()
  Dim FD As FileDialog
  Set FD = Application.FileDialog(msoFileDialogOpen)
  With FD
      .Filters.add "src files", "*.xls; *.xlsx; *.xlsm; *.csv", 1
      'change default directory
      .InitialFileName = "C:\Users\" & Environ("USERNAME") & "\Downloads\"
      
      If .Show = True Then ' if valid
        ' add Selected items to filePathList
        Dim it
        For Each it In .SelectedItems
          filePathList.add it
        Next
      Else ' if INvalid operation
          'cancel button pressed
      End If
    End With
End Sub

Public Sub showAllPath()
  Dim fp
  For Each fp In filePathList
     Debug.Print fp
  Next
End Sub

Public Function nextItem()
  nextItem = filePathList(index)
  index = index + 1
End Function

Public Function hasNext() As Boolean
  hasNext = index <= filePathList.Count
End Function
