Option Explicit

' ReadRAM es Read(address) del enunciado.
' WriteRAM es Write(address, value) del enunciado.
' La hoja RAM solo cambia dentro de WriteRAM.

Private Function HojaRAM() As Worksheet
    Set HojaRAM = ThisWorkbook.Worksheets("RAM")
End Function

Private Function HojaMem() As Worksheet
    Set HojaMem = ThisWorkbook.Worksheets("Memoria")
End Function

Private Function HojaCPU() As Worksheet
    On Error Resume Next
    Set HojaCPU = ThisWorkbook.Worksheets("CPU")
End Function

Private Sub PonerMarMdr(ByVal address As Long, ByVal value As Long)
    Dim cpu As Worksheet
    HojaMem.Range("F4").Value = address
    HojaMem.Range("F5").Value = value
    Set cpu = HojaCPU()
    If cpu Is Nothing Then Exit Sub
    cpu.Range("B8").Value = address
    cpu.Range("B9").Value = value
End Sub

Public Function PeekRAM(ByVal address As Long) As Long
    Dim v As Variant
    If address < 0 Or address > 255 Then
        PeekRAM = 0
        Exit Function
    End If
    v = HojaRAM.Cells(address + 1, 1).Value
    If IsNumeric(v) Then
        PeekRAM = CLng(v) And 255
    Else
        PeekRAM = 0
    End If
End Function

Public Function ReadRAM(ByVal address As Long) As Long
    Dim v As Long
    If address < 0 Or address > 255 Then
        MsgBox "La casilla debe estar entre 0 y 255."
        ReadRAM = 0
        Exit Function
    End If
    v = PeekRAM(address)
    PonerMarMdr address, v
    PintarCasilla address
    ReadRAM = v
End Function

Public Sub WriteRAM(ByVal address As Long, ByVal value As Long)
    If address < 0 Or address > 255 Then
        MsgBox "La casilla debe estar entre 0 y 255."
        Exit Sub
    End If
    If value < 0 Or value > 255 Then
        MsgBox "El valor debe estar entre 0 y 255."
        Exit Sub
    End If
    HojaRAM.Cells(address + 1, 1).Value = value
    PonerMarMdr address, value
    PintarCasilla address
End Sub

Public Sub BotonLeer()
    Dim addr As Long
    Dim v As Long
    addr = NumeroEn(HojaMem.Range("B4"), 0, 255, "Casilla")
    If addr < 0 Then Exit Sub
    v = ReadRAM(addr)
    HojaMem.Range("B5").Value = v
    HojaMem.Range("A7").Value = "Leida la casilla " & addr & ". Vale " & v & "."
End Sub

Public Sub BotonEscribir()
    Dim addr As Long
    Dim v As Long
    addr = NumeroEn(HojaMem.Range("B4"), 0, 255, "Casilla")
    If addr < 0 Then Exit Sub
    v = NumeroEn(HojaMem.Range("B5"), 0, 255, "Valor")
    If v < 0 Then Exit Sub
    WriteRAM addr, v
    HojaMem.Range("A7").Value = "Guardado " & v & " en la casilla " & addr & "."
End Sub

Public Sub BotonProbar()
    Dim v As Long
    HojaMem.Range("B4").Value = 128
    HojaMem.Range("B5").Value = 42
    WriteRAM 128, 42
    v = ReadRAM(128)
    If v = 42 Then
        HojaMem.Range("A7").Value = "Prueba bien: la casilla 128 vale 42. En el mapa se ve 2A, 00101010 y 42."
    Else
        HojaMem.Range("A7").Value = "Prueba mal."
    End If
End Sub

Public Sub BotonLimpiar()
    Dim i As Long
    For i = 1 To 256
        HojaRAM.Cells(i, 1).Value = 0
    Next i
    PonerMarMdr 0, 0
    HojaMem.Range("B4").Value = 0
    HojaMem.Range("B5").Value = 0
    HojaMem.Range("Z1").Value = -1
    RepintarMapa
    HojaMem.Range("A7").Value = "Memoria en cero."
End Sub

Private Function NumeroEn(ByVal celda As Range, ByVal minimo As Long, ByVal maximo As Long, ByVal nombre As String) As Long
    If Len(Trim$(CStr(celda.Value))) = 0 Or Not IsNumeric(celda.Value) Then
        MsgBox nombre & " tiene que ser un numero."
        NumeroEn = -1
        Exit Function
    End If
    If CLng(celda.Value) < minimo Or CLng(celda.Value) > maximo Then
        MsgBox nombre & " tiene que estar entre " & minimo & " y " & maximo & "."
        NumeroEn = -1
        Exit Function
    End If
    NumeroEn = CLng(celda.Value)
End Function

Public Sub PintarCasilla(ByVal address As Long)
    Dim ws As Worksheet
    Dim prev As Long
    Set ws = HojaMem
    If IsNumeric(ws.Range("Z1").Value) Then
        prev = CLng(ws.Range("Z1").Value)
        If prev >= 0 And prev <= 255 And prev <> address Then
            ColorBase CeldaDe(prev), prev
        End If
    End If
    CeldaDe(address).Interior.Color = RGB(249, 231, 159)
    ws.Range("Z1").Value = address
End Sub

Public Sub RepintarMapa()
    Dim addr As Long
    For addr = 0 To 255
        ColorBase CeldaDe(addr), addr
    Next addr
End Sub

Private Function CeldaDe(ByVal address As Long) As Range
    Set CeldaDe = HojaMem.Cells(11 + (address \ 16), 2 + (address Mod 16))
End Function

Private Sub ColorBase(ByVal celda As Range, ByVal address As Long)
    If address < 128 Then
        celda.Interior.Color = RGB(214, 234, 248)
    Else
        celda.Interior.Color = RGB(213, 245, 227)
    End If
End Sub
