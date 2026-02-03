Sub FixTextFormulasInWorkbook_v2()
  Dim ws As Worksheet, c As Range, rng As Range
  Dim cnt As Long
  Dim s As String

  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  cnt = 0

  For Each ws In ThisWorkbook.Worksheets
    On Error Resume Next
    Set rng = ws.UsedRange
    On Error GoTo 0

    If Not rng Is Nothing Then
      ' Força formato Geral (remove formatação Texto)
      rng.NumberFormat = "General"

      For Each c In rng.Cells
        If Not IsError(c.Value2) Then
          s = CStr(c.Value2)
          ' remover espaços normais, NBSP e zero-width space
          s = Trim(Replace(s, Chr(160), ""))
          s = Replace(s, ChrW(&H200B), "")

          If Len(s) > 0 Then
            ' Caso: a célula contém exatamente uma string que começa com "="
            If Left(s, 1) = "=" Then
              c.Formula = s
              cnt = cnt + 1

            ' Caso: a célula contém "'=..." (apóstrofo + igual)
            ElseIf Len(s) >= 2 And Left(s, 2) = "'=" Then
              c.Formula = Mid(s, 2)
              cnt = cnt + 1

            ' Caso extra: a propriedade .Formula começa com apóstrofo (algumas versões)
            ElseIf Len(CStr(c.Formula)) > 0 And Left(CStr(c.Formula), 1) = "'" Then
              c.Formula = Mid(CStr(c.Formula), 2)
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
