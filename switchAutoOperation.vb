Option Explicit

Sub disableAutoOperationBeforeExecMacro()
  Application.DisplayAlerts = False
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
End Sub

Sub enableAutoOperationAfterExecMacro()
  Application.DisplayAlerts = True
  Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
End Sub
