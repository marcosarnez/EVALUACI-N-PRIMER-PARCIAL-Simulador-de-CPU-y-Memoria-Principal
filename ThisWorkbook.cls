Private Sub Workbook_Open()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect ""
    Next ws
    ThisWorkbook.Worksheets("Memoria").Visible = 0
    ThisWorkbook.Worksheets("Pruebas").Visible = 0
    ThisWorkbook.Worksheets("Estado").Visible = 2
    ThisWorkbook.Worksheets("CPU").Protect "", False, True, True, True
    ThisWorkbook.Worksheets("Memoria").Protect "", False, True, True, True
    ThisWorkbook.Worksheets("Pruebas").Protect "", False, True, True, True
    ThisWorkbook.Worksheets("RAM").Protect "", True, True, True, True
    ThisWorkbook.Worksheets("Inspeccion").Protect "", True, True, True, True
    ThisWorkbook.Worksheets("Estado").Protect "", True, True, True, True
    ThisWorkbook.Worksheets("CPU").Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
End Sub
