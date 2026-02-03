Attribute VB_Name = "FixFormulasModule"
Sub FixTextFormulasInWorkbook()
  Dim ws As Worksheet, c As Range, rng As Range
  Dim cnt As Long
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  cnt = 0
  For Each ws In ThisWorkbook.Worksheets
    On Error Resume Next
    Set rng = ws.UsedRange
    On Error GoTo 0
    If Not rng Is Nothing Then
      rng.NumberFormat = "General"
      For Each c In rng.Cells
        If Len(c.Text) > 0 Then
          If Left(c.Formula, 1) = "'" And Len(c.Formula) > 1 Then
            c.Formula = Mid(c.Formula, 2)
            cnt = cnt + 1
          Else
            If Left(c.Text, 1) = "=" Then
              c.Formula = c.Text
              cnt = cnt + 1
            End If
          End If
        End If
      Next c
    End If
    Set rng = Nothing
  Next ws
  Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  MsgBox "Corrigidas " & cnt & " células que estavam como texto e continham fórmulas.", vbInformation
End Sub