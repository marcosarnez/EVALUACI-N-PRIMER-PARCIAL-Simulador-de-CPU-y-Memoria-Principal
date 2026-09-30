Option Explicit

' Unica rutina que escribe en la hoja CPU.
' Registros, banderas, fases, mapa, log y estado usan nombres, no celdas sueltas.

Private Function Nom(ByVal nombre As String) As Range
    Set Nom = ThisWorkbook.Names(nombre).RefersToRange
End Function

Private Function HojaEstado() As Worksheet
    Set HojaEstado = ThisWorkbook.Worksheets("Estado")
End Function

Private Function EstadoCelda(ByVal clave As String) As Range
    Dim n As Long
    If clave = "MapPrev" Then
        Set EstadoCelda = HojaEstado.Cells(50, 1)
    Else
        n = CLng(Mid$(clave, 2))
        Set EstadoCelda = HojaEstado.Cells(n, 1)
    End If
End Function

Public Function UiState(ByVal clave As String) As Long
    Dim v As Variant
    v = EstadoCelda(clave).Value
    If IsNumeric(v) Then
        UiState = CLng(v)
    Else
        UiState = 0
    End If
End Function

Public Function UiStateText(ByVal clave As String) As String
    Dim v As Variant
    v = EstadoCelda(clave).Value
    If IsEmpty(v) Then
        UiStateText = ""
    Else
        UiStateText = CStr(v)
    End If
End Function

Public Sub UiStateSet(ByVal clave As String, ByVal valor As Variant)
    EstadoCelda(clave).Value = valor
End Sub

Public Function UiGetReg(ByVal nombre As String) As Long
    Dim v As Variant
    v = 0
    Select Case UCase$(nombre)
        Case "PC": v = Nom("PC_Dec").Value
        Case "IR": v = Nom("IR_Dec").Value
        Case "MAR": v = Nom("MAR_Dec").Value
        Case "MDR": v = Nom("MDR_Dec").Value
        Case "AX": v = Nom("AX_Dec").Value
        Case "BX": v = Nom("BX_Dec").Value
        Case "ZF": v = Nom("Flag_ZF").Value
        Case "CF": v = Nom("Flag_CF").Value
        Case "SF": v = Nom("Flag_SF").Value
    End Select
    If IsNumeric(v) Then
        UiGetReg = CLng(v) And 255
    Else
        UiGetReg = 0
    End If
End Function

Public Sub UiSetReg(ByVal nombre As String, ByVal valor As Long)
    valor = valor And 255
    Select Case UCase$(nombre)
        Case "PC": Nom("PC_Dec").Value = valor
        Case "IR": Nom("IR_Dec").Value = valor
        Case "MAR": Nom("MAR_Dec").Value = valor
        Case "MDR": Nom("MDR_Dec").Value = valor
        Case "AX": Nom("AX_Dec").Value = valor
        Case "BX": Nom("BX_Dec").Value = valor
        Case "ZF": Nom("Flag_ZF").Value = IIf(valor <> 0, 1, 0)
        Case "CF": Nom("Flag_CF").Value = IIf(valor <> 0, 1, 0)
        Case "SF": Nom("Flag_SF").Value = IIf(valor <> 0, 1, 0)
    End Select
    If UCase$(nombre) = "ZF" Or UCase$(nombre) = "CF" Or UCase$(nombre) = "SF" Then
        UiPaintFlags
    End If
End Sub

Public Sub UiPaintFlags()
    PintarFlag Nom("Flag_ZF"), UiGetReg("ZF")
    PintarFlag Nom("Flag_CF"), UiGetReg("CF")
    PintarFlag Nom("Flag_SF"), UiGetReg("SF")
End Sub

Private Sub PintarFlag(ByVal celda As Range, ByVal bit As Long)
    celda.Font.Name = "Consolas"
    celda.Font.Bold = True
    celda.HorizontalAlignment = -4108
    If bit = 1 Then
        celda.Interior.Color = RGB(249, 231, 159)
        celda.Font.Color = RGB(31, 78, 121)
    Else
        celda.Interior.Color = RGB(255, 255, 255)
        celda.Font.Color = RGB(40, 40, 40)
    End If
End Sub

Public Sub UiStatus(ByVal texto As String)
    Nom("Status_Msg").Value = texto
End Sub

Public Sub UiMnemonic(ByVal texto As String)
    Nom("Sel_Mnem").Value = texto
End Sub

Public Function UiMnemText() As String
    Dim v As Variant
    v = Nom("Sel_Mnem").Value
    If IsEmpty(v) Then
        UiMnemText = ""
    Else
        UiMnemText = CStr(v)
    End If
End Function

Public Function UiDelay() As Long
    Dim v As Variant
    v = Nom("Delay_Ms").Value
    If IsNumeric(v) Then
        UiDelay = CLng(v)
    Else
        UiDelay = 250
    End If
    If UiDelay < 0 Then UiDelay = 0
    If UiDelay > 2000 Then UiDelay = 2000
End Function

Public Sub UiLog(ByVal texto As String)
    Dim i As Long
    For i = 10 To 2 Step -1
        Nom("Log_" & CStr(i)).Value = Nom("Log_" & CStr(i - 1)).Value
    Next i
    Nom("Log_1").Value = texto
End Sub

Public Sub UiClearLog()
    Dim i As Long
    For i = 1 To 10
        Nom("Log_" & CStr(i)).Value = ""
    Next i
End Sub

Public Sub UiPintarFase(ByVal activa As Long)
    PintarFase Nom("Phase_Fetch"), activa = 0
    PintarFase Nom("Phase_Decode"), activa = 1
    PintarFase Nom("Phase_Execute"), activa = 2
    PintarFase Nom("Phase_Store"), activa = 3
    UiPintarRegs activa
End Sub

Private Sub PintarFase(ByVal celda As Range, ByVal onOff As Boolean)
    celda.Font.Name = "Calibri"
    celda.Font.Bold = True
    celda.HorizontalAlignment = -4108
    If onOff Then
        celda.Interior.Color = RGB(31, 78, 121)
        celda.Font.Color = RGB(255, 255, 255)
    Else
        celda.Interior.Color = RGB(220, 228, 236)
        celda.Font.Color = RGB(31, 78, 121)
    End If
End Sub

Public Sub UiPintarRegs(ByVal fase As Long)
    Dim nombres As Variant
    Dim i As Long
    Dim activo As Boolean
    Dim dec As Range
    nombres = Array("PC_Dec", "IR_Dec", "MAR_Dec", "MDR_Dec", "AX_Dec", "BX_Dec")
    For i = 0 To 5
        activo = False
        If fase = 0 And i <= 3 Then activo = True
        If fase = 1 And i = 1 Then activo = True
        If fase = 2 And (i = 4 Or i = 5) Then activo = True
        If fase = 3 And (i = 2 Or i = 3 Or i = 4) Then activo = True
        Set dec = Nom(CStr(nombres(i)))
        If activo Then
            dec.Offset(0, 3).Value = "ACTIVO"
            dec.Worksheet.Range(dec.Offset(0, -1), dec.Offset(0, 3)).Interior.Color = RGB(249, 231, 159)
        Else
            dec.Offset(0, 3).Value = ""
            dec.Worksheet.Range(dec.Offset(0, -1), dec.Offset(0, 3)).Interior.Color = RGB(255, 255, 255)
        End If
        dec.Font.Name = "Consolas"
        dec.Font.Bold = True
        dec.Offset(0, 3).Font.Name = "Calibri"
        dec.Offset(0, 3).Font.Bold = True
        dec.Offset(0, 3).HorizontalAlignment = -4108
    Next i
End Sub

Private Function CeldaMapa(ByVal address As Long) As Range
    ' Fila 19, columna I. No comparte filas con el panel de registros.
    Set CeldaMapa = ThisWorkbook.Worksheets("CPU").Cells(19 + (address \ 16), 9 + (address Mod 16))
End Function

Private Sub ColorBaseMapa(ByVal address As Long)
    Dim celda As Range
    Set celda = CeldaMapa(address)
    If address < 128 Then
        celda.Interior.Color = RGB(214, 234, 248)
    Else
        celda.Interior.Color = RGB(213, 245, 227)
    End If
End Sub

Public Sub UiPintarMapa(ByVal address As Long)
    Dim prev As Long
    If address < 0 Or address > 255 Then Exit Sub
    prev = UiState("MapPrev")
    If prev >= 0 And prev <= 255 And prev <> address Then
        ColorBaseMapa prev
    End If
    CeldaMapa(address).Interior.Color = RGB(249, 231, 159)
    UiStateSet "MapPrev", address
End Sub

Public Sub UiRepintarMapa()
    Dim addr As Long
    For addr = 0 To 255
        ColorBaseMapa addr
    Next addr
    UiStateSet "MapPrev", -1
End Sub
