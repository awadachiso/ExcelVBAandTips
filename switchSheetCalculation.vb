Option Explicit

Private prevSheet As Worksheet, curSheet As Worksheet


Private Sub Workbook_Open()
  Application.Calculation = xlCalculationManual
  
  Dim s As Worksheet
  'set all sheets disable Calclation
  For Each s In Worksheets
    s.EnableCalculation = False
  Next
  
  Set curSheet = ActiveSheet
End Sub

'Event select another sheet
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
  Set prevSheet = curSheet
  Set curSheet = ActiveSheet
  
  prevSheet.EnableCalculation = False
  curSheet.EnableCalculation = True
  curSheet.Calculate  
End Sub
