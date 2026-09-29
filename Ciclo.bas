Option Explicit

' Ciclo de instruccion: Fetch -> Decode -> Execute -> Store.
' Un STEP avanza una sola fase.

Private Function HojaCPU() As Worksheet
    Set HojaCPU = ThisWorkbook.Worksheets("CPU")
End Function

Private Function NzL(ByVal celda As String) As Long
    Dim v As Variant
    v = HojaCPU.Range(celda).Value
    If IsNumeric(v) Then
        NzL = CLng(v)
    Else
        NzL = 0
    End If
End Function

Private Function Hx(ByVal n As Long) As String
    Hx = Right$("0" & Hex$(n And 255), 2) & "h"
End Function

Private Function NombreReg(ByVal n As Long) As String
    If (n And 255) = 1 Then
        NombreReg = "BX"
    Else
        NombreReg = "AX"
    End If
End Function

Private Function ValorReg(ByVal n As Long) As Long
    If (n And 255) = 1 Then
        ValorReg = GetReg("BX")
    Else
        ValorReg = GetReg("AX")
    End If
End Function

Private Sub DestinoReg(ByVal n As Long)
    If (n And 255) = 1 Then
        HojaCPU.Range("Z6").Value = 2
    Else
        HojaCPU.Range("Z6").Value = 1
    End If
End Sub

Private Function TamanoOpcode(ByVal op As Long) As Long
    Select Case op
        Case &H0
            TamanoOpcode = 1
        Case &H14, &H15, &H1B, &H20, &H21, &H22, &H23
            TamanoOpcode = 2
        Case Else
            TamanoOpcode = 3
    End Select
End Function

Private Sub LogLine(ByVal texto As String)
    Dim i As Long
    For i = 45 To 35 Step -1
        HojaCPU.Cells(i, 1).Value = HojaCPU.Cells(i - 1, 1).Value
    Next i
    HojaCPU.Range("A34").Value = texto
    HojaCPU.Range("A22").Value = texto
End Sub

Private Sub PintarFases(ByVal activa As Long)
    PintarCaja "B17", activa = 0, RGB(61, 126, 166)
    PintarCaja "D17", activa = 1, RGB(138, 109, 59)
    PintarCaja "F17", activa = 2, RGB(74, 124, 78)
    PintarCaja "H17", activa = 3, RGB(122, 78, 138)
End Sub

Private Sub PintarCaja(ByVal celda As String, ByVal onOff As Boolean, ByVal colorOn As Long)
    If onOff Then
        HojaCPU.Range(celda).Interior.Color = colorOn
        HojaCPU.Range(celda).Font.Color = RGB(255, 255, 255)
    Else
        HojaCPU.Range(celda).Interior.Color = RGB(200, 205, 210)
        HojaCPU.Range(celda).Font.Color = RGB(50, 50, 50)
    End If
End Sub

Private Sub LimpiarEstadoCiclo()
    Dim i As Long
    HojaCPU.Range("Z1").Value = 0
    HojaCPU.Range("Z2").Value = 0
    For i = 3 To 21
        HojaCPU.Range("Z" & i).Value = 0
    Next i
    HojaCPU.Range("Z9").Value = 0
    HojaCPU.Range("Z30").Value = ""
    HojaCPU.Range("A19").Value = "Pulsa CARGAR PROGRAMA y luego STEP."
    PintarFases -1
End Sub

Public Sub BotonResetCiclo()
    Dim i As Long
    BotonResetCPU
    LimpiarEstadoCiclo
    For i = 34 To 45
        HojaCPU.Cells(i, 1).Value = ""
    Next i
    HojaCPU.Range("A22").Value = "RESET: registros y PC en cero. La memoria no se borra."
End Sub

Public Sub BotonCargarPrograma()
    Dim i As Long
    BotonResetCiclo
    For i = 0 To 15
        WriteRAM i, 0
    Next i
    WriteRAM 0, &H3
    WriteRAM 1, 0
    WriteRAM 2, &H80
    WriteRAM 3, &H3
    WriteRAM 4, 1
    WriteRAM 5, &H81
    WriteRAM 6, &H11
    WriteRAM 7, 0
    WriteRAM 8, 1
    WriteRAM 9, &H4
    WriteRAM 10, 0
    WriteRAM 11, &H82
    WriteRAM 12, 0
    WriteRAM &H80, 5
    WriteRAM &H81, 3
    WriteRAM &H82, 0
    SetReg "PC", 0
    SetReg "IR", 0
    SetReg "MAR", 0
    SetReg "MDR", 0
    PintarCasilla 0
    PintarFases -1
    LogLine "Programa: LOAD AX,[80h]  LOAD BX,[81h]  ADD AX,BX  STORE [82h],AX  HLT. Datos 5 y 3."
End Sub

Public Sub BotonSTEP()
    If NzL("Z2") = 1 Then
        HojaCPU.Range("A22").Value = "Detenido con HLT. Pulsa RESET o CARGAR PROGRAMA."
        PintarFases -1
        Exit Sub
    End If
    HojaCPU.Range("Z9").Value = NzL("Z9") + 1
    Select Case NzL("Z1")
        Case 0
            HacerFetch
        Case 1
            HacerDecode
        Case 2
            HacerExecute
        Case Else
            HacerStore
    End Select
End Sub

Private Sub HacerFetch()
    Dim pc0 As Long
    Dim op As Long
    Dim sz As Long
    pc0 = GetReg("PC")
    op = ReadRAM(pc0)
    SetReg "IR", op
    sz = TamanoOpcode(op)
    If sz >= 2 Then
        HojaCPU.Range("Z3").Value = PeekRAM((pc0 + 1) And 255)
    Else
        HojaCPU.Range("Z3").Value = 0
    End If
    If sz >= 3 Then
        HojaCPU.Range("Z4").Value = PeekRAM((pc0 + 2) And 255)
    Else
        HojaCPU.Range("Z4").Value = 0
    End If
    HojaCPU.Range("Z8").Value = sz
    SetReg "PC", (pc0 + sz) And 255
    HojaCPU.Range("Z1").Value = 1
    PintarFases 0
    LogLine "[Paso " & Format$(NzL("Z9"), "00") & "] FETCH: PC=" & Hx(pc0) & " MAR=" & Hx(pc0) & " MDR=" & Hx(op) & " IR=" & Hx(op) & " PC<-" & Hx((pc0 + sz) And 255)
End Sub

Private Sub HacerDecode()
    Dim op As Long
    Dim a As Long
    Dim b As Long
    Dim texto As String
    op = GetReg("IR")
    a = NzL("Z3")
    b = NzL("Z4")
    HojaCPU.Range("Z6").Value = 0
    HojaCPU.Range("Z7").Value = 0
    HojaCPU.Range("Z16").Value = 0
    HojaCPU.Range("Z17").Value = 0
    HojaCPU.Range("Z19").Value = 0
    HojaCPU.Range("Z20").Value = 0
    HojaCPU.Range("Z21").Value = 0
    HojaCPU.Range("Z30").Value = ""
    texto = "opcode desconocido " & Hx(op)
    Select Case op
        Case &H0
            texto = "HLT"
            HojaCPU.Range("Z20").Value = 1
            HojaCPU.Range("Z30").Value = "HLT"
        Case &H1
            texto = "MOV " & NombreReg(a) & ", " & Hx(b)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "MOVIMM"
        Case &H2
            texto = "MOV " & NombreReg(a) & ", " & NombreReg(b)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "MOVREG"
        Case &H3
            texto = "LOAD " & NombreReg(a) & ", [" & Hx(b) & "]"
            DestinoReg a
            HojaCPU.Range("Z7").Value = b
            HojaCPU.Range("Z30").Value = "LOAD"
        Case &H4
            texto = "STORE [" & Hx(b) & "], " & NombreReg(a)
            HojaCPU.Range("Z6").Value = 3
            HojaCPU.Range("Z7").Value = b
            HojaCPU.Range("Z16").Value = 1
            HojaCPU.Range("Z30").Value = "STORE"
        Case &H10
            texto = "ADD " & NombreReg(a) & ", " & Hx(b)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "ADDIMM"
        Case &H11
            texto = "ADD " & NombreReg(a) & ", " & NombreReg(b)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "ADDREG"
        Case &H12
            texto = "SUB " & NombreReg(a) & ", " & Hx(b)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "SUBIMM"
        Case &H13
            texto = "SUB " & NombreReg(a) & ", " & NombreReg(b)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "SUBREG"
        Case &H14
            texto = "INC " & NombreReg(a)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "INC"
        Case &H15
            texto = "DEC " & NombreReg(a)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "DEC"
        Case &H16
            texto = "CMP " & NombreReg(a) & ", " & Hx(b)
            HojaCPU.Range("Z21").Value = 1
            HojaCPU.Range("Z30").Value = "CMPIMM"
        Case &H17
            texto = "CMP " & NombreReg(a) & ", " & NombreReg(b)
            HojaCPU.Range("Z21").Value = 1
            HojaCPU.Range("Z30").Value = "CMPREG"
        Case &H18
            texto = "AND " & NombreReg(a) & ", " & NombreReg(b)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "AND"
        Case &H19
            texto = "OR " & NombreReg(a) & ", " & NombreReg(b)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "OR"
        Case &H1A
            texto = "XOR " & NombreReg(a) & ", " & NombreReg(b)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "XOR"
        Case &H1B
            texto = "NOT " & NombreReg(a)
            DestinoReg a
            HojaCPU.Range("Z30").Value = "NOT"
        Case &H20
            texto = "JMP " & Hx(a)
            HojaCPU.Range("Z17").Value = 1
            HojaCPU.Range("Z18").Value = a
            HojaCPU.Range("Z19").Value = 0
            HojaCPU.Range("Z30").Value = "JMP"
        Case &H21
            texto = "JZ " & Hx(a)
            HojaCPU.Range("Z17").Value = 1
            HojaCPU.Range("Z18").Value = a
            HojaCPU.Range("Z19").Value = 1
            HojaCPU.Range("Z30").Value = "JZ"
        Case &H22
            texto = "JNZ " & Hx(a)
            HojaCPU.Range("Z17").Value = 1
            HojaCPU.Range("Z18").Value = a
            HojaCPU.Range("Z19").Value = 2
            HojaCPU.Range("Z30").Value = "JNZ"
        Case &H23
            texto = "JC " & Hx(a)
            HojaCPU.Range("Z17").Value = 1
            HojaCPU.Range("Z18").Value = a
            HojaCPU.Range("Z19").Value = 3
            HojaCPU.Range("Z30").Value = "JC"
    End Select
    HojaCPU.Range("A19").Value = texto
    HojaCPU.Range("Z1").Value = 2
    PintarFases 1
    LogLine "[Paso " & Format$(NzL("Z9"), "00") & "] DECODE: " & texto
End Sub

Private Sub HacerExecute()
    Dim kind As String
    Dim a As Long
    Dim b As Long
    Dim r As Long
    Dim toma As Boolean
    kind = CStr(HojaCPU.Range("Z30").Value)
    a = NzL("Z3")
    b = NzL("Z4")
    r = 0
    Select Case kind
        Case "MOVIMM"
            r = b
        Case "MOVREG"
            r = ValorReg(b)
        Case "LOAD"
            r = ReadRAM(NzL("Z7"))
        Case "STORE"
            r = ValorReg(a)
            PonerMarMdrVisible NzL("Z7"), r
        Case "ADDIMM"
            r = ALU("ADD", ValorReg(a), b)
        Case "ADDREG"
            r = ALU("ADD", ValorReg(a), ValorReg(b))
        Case "SUBIMM"
            r = ALU("SUB", ValorReg(a), b)
        Case "SUBREG"
            r = ALU("SUB", ValorReg(a), ValorReg(b))
        Case "INC"
            r = ALU("INC", ValorReg(a), 1)
        Case "DEC"
            r = ALU("DEC", ValorReg(a), 1)
        Case "CMPIMM"
            r = ALU("CMP", ValorReg(a), b)
        Case "CMPREG"
            r = ALU("CMP", ValorReg(a), ValorReg(b))
        Case "AND"
            r = ALU("AND", ValorReg(a), ValorReg(b))
        Case "OR"
            r = ALU("OR", ValorReg(a), ValorReg(b))
        Case "XOR"
            r = ALU("XOR", ValorReg(a), ValorReg(b))
        Case "NOT"
            r = ALU("NOT", ValorReg(a), 0)
        Case "JMP", "JZ", "JNZ", "JC"
            toma = False
            If NzL("Z19") = 0 Then toma = True
            If NzL("Z19") = 1 And GetReg("ZF") = 1 Then toma = True
            If NzL("Z19") = 2 And GetReg("ZF") = 0 Then toma = True
            If NzL("Z19") = 3 And GetReg("CF") = 1 Then toma = True
            If toma Then
                SetReg "PC", NzL("Z18")
            End If
        Case "HLT"
            r = 0
    End Select
    HojaCPU.Range("Z5").Value = r
    HojaCPU.Range("Z1").Value = 3
    PintarFases 2
    LogLine "[Paso " & Format$(NzL("Z9"), "00") & "] EXECUTE: " & CStr(HojaCPU.Range("A19").Value) & "  resultado=" & r
End Sub

Private Sub PonerMarMdrVisible(ByVal address As Long, ByVal value As Long)
    Dim dummy As Long
    dummy = PeekRAM(address)
    ThisWorkbook.Worksheets("Memoria").Range("F4").Value = address
    ThisWorkbook.Worksheets("Memoria").Range("F5").Value = value
    HojaCPU.Range("B8").Value = address
    HojaCPU.Range("B9").Value = value
    PintarCasilla address
End Sub

Private Sub HacerStore()
    Dim dest As Long
    Dim r As Long
    Dim kind As String
    dest = NzL("Z6")
    r = NzL("Z5")
    kind = CStr(HojaCPU.Range("Z30").Value)
    If NzL("Z21") = 0 And dest = 1 Then SetReg "AX", r
    If NzL("Z21") = 0 And dest = 2 Then SetReg "BX", r
    If dest = 3 Then WriteRAM NzL("Z7"), r
    If NzL("Z20") = 1 Then
        HojaCPU.Range("Z2").Value = 1
        LogLine "[Paso " & Format$(NzL("Z9"), "00") & "] STORE: HLT. El reloj se detiene."
    ElseIf dest = 3 Then
        LogLine "[Paso " & Format$(NzL("Z9"), "00") & "] STORE: RAM[" & Hx(NzL("Z7")) & "] <- " & r
    ElseIf dest = 1 Then
        LogLine "[Paso " & Format$(NzL("Z9"), "00") & "] STORE: AX <- " & r
    ElseIf dest = 2 Then
        LogLine "[Paso " & Format$(NzL("Z9"), "00") & "] STORE: BX <- " & r
    Else
        LogLine "[Paso " & Format$(NzL("Z9"), "00") & "] STORE: sin escritura. PC=" & Hx(GetReg("PC"))
    End If
    HojaCPU.Range("Z1").Value = 0
    PintarFases 3
End Sub
