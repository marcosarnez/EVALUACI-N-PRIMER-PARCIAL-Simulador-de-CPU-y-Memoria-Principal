Private Sub Workbook_Open()
    On Error Resume Next
    ThisWorkbook.Worksheets("CPU").Unprotect
    ThisWorkbook.Worksheets("Memoria").Unprotect
    ThisWorkbook.Worksheets("RAM").Unprotect
    ThisWorkbook.Worksheets("Inspeccion").Unprotect
    ThisWorkbook.Worksheets("CPU").Protect "", False, True, True, True
    ThisWorkbook.Worksheets("Memoria").Protect "", False, True, True, True
    ThisWorkbook.Worksheets("RAM").Protect "", True, True, True, True
    ThisWorkbook.Worksheets("Inspeccion").Protect "", True, True, True, True
End Sub
