Option Explicit

' Banco de registros y ALU de 8 bits.
' AX y BX se pueden escribir en la hoja. El resto los mueve el simulador.

Private Function HojaCPU() As Worksheet
    Set HojaCPU = ThisWorkbook.Worksheets("CPU")
End Function

Private Function Byte8(ByVal valor As Long) As Long
    Byte8 = valor And 255
End Function

Public Function GetReg(ByVal nombre As String) As Long
    Dim v As Variant
    v = 0
    Select Case UCase$(nombre)
        Case "PC": v = HojaCPU.Range("B6").Value
        Case "IR": v = HojaCPU.Range("B7").Value
        Case "MAR": v = HojaCPU.Range("B8").Value
        Case "MDR": v = HojaCPU.Range("B9").Value
        Case "AX": v = HojaCPU.Range("B10").Value
        Case "BX": v = HojaCPU.Range("B11").Value
        Case "ZF": v = HojaCPU.Range("E13").Value
        Case "CF": v = HojaCPU.Range("G13").Value
        Case "SF": v = HojaCPU.Range("I13").Value
    End Select
    If IsNumeric(v) Then
        GetReg = CLng(v)
    Else
        GetReg = 0
    End If
End Function

Public Sub SetReg(ByVal nombre As String, ByVal valor As Long)
    valor = Byte8(valor)
    Select Case UCase$(nombre)
        Case "PC": HojaCPU.Range("B6").Value = valor
        Case "IR": HojaCPU.Range("B7").Value = valor
        Case "MAR": HojaCPU.Range("B8").Value = valor
        Case "MDR": HojaCPU.Range("B9").Value = valor
        Case "AX": HojaCPU.Range("B10").Value = valor
        Case "BX": HojaCPU.Range("B11").Value = valor
        Case "ZF": HojaCPU.Range("E13").Value = IIf(valor <> 0, 1, 0)
        Case "CF": HojaCPU.Range("G13").Value = IIf(valor <> 0, 1, 0)
        Case "SF": HojaCPU.Range("I13").Value = IIf(valor <> 0, 1, 0)
    End Select
    PintarFlags
End Sub

Public Sub PintarFlags()
    PintarFlag "E13", GetReg("ZF")
    PintarFlag "G13", GetReg("CF")
    PintarFlag "I13", GetReg("SF")
End Sub

Private Sub PintarFlag(ByVal celda As String, ByVal bit As Long)
    If bit = 1 Then
        HojaCPU.Range(celda).Interior.Color = RGB(46, 204, 113)
        HojaCPU.Range(celda).Font.Color = RGB(20, 40, 20)
    Else
        HojaCPU.Range(celda).Interior.Color = RGB(189, 195, 199)
        HojaCPU.Range(celda).Font.Color = RGB(40, 40, 40)
    End If
End Sub

' Calcula en 8 bits y actualiza ZF, CF y SF.
' CMP usa la resta pero no guarda el resultado.
Public Function ALU(ByVal op As String, ByVal a As Long, ByVal b As Long) As Long
    Dim r As Long
    Dim c As Long
    Dim logica As Boolean
    a = Byte8(a)
    b = Byte8(b)
    c = 0
    logica = False
    Select Case UCase$(op)
        Case "ADD"
            r = a + b
            If r > 255 Then c = 1
            r = Byte8(r)
        Case "SUB", "CMP"
            If a < b Then
                c = 1
                r = Byte8(a - b + 256)
            Else
                r = a - b
            End If
        Case "INC"
            r = a + 1
            If r > 255 Then c = 1
            r = Byte8(r)
        Case "DEC"
            If a = 0 Then
                c = 1
                r = 255
            Else
                r = a - 1
            End If
        Case "AND"
            r = a And b
            logica = True
        Case "OR"
            r = a Or b
            logica = True
        Case "XOR"
            r = a Xor b
            logica = True
        Case "NOT"
            r = (Not a) And 255
            logica = True
        Case Else
            r = a
    End Select
    HojaCPU.Range("E13").Value = IIf(r = 0, 1, 0)
    If logica Then
        HojaCPU.Range("G13").Value = 0
    Else
        HojaCPU.Range("G13").Value = c
    End If
    HojaCPU.Range("I13").Value = IIf((r And 128) <> 0, 1, 0)
    PintarFlags
    ALU = r
End Function

Private Sub Mensaje(ByVal texto As String)
    HojaCPU.Range("A22").Value = texto
End Sub

Public Sub BotonADD()
    Dim ax As Long
    Dim bx As Long
    Dim r As Long
    ax = GetReg("AX")
    bx = GetReg("BX")
    r = ALU("ADD", ax, bx)
    SetReg "AX", r
    Mensaje "ADD: AX = " & ax & " + " & bx & " = " & r & ". ZF=" & GetReg("ZF") & " CF=" & GetReg("CF") & " SF=" & GetReg("SF")
End Sub

Public Sub BotonSUB()
    Dim ax As Long
    Dim bx As Long
    Dim r As Long
    ax = GetReg("AX")
    bx = GetReg("BX")
    r = ALU("SUB", ax, bx)
    SetReg "AX", r
    Mensaje "SUB: AX = " & ax & " - " & bx & " = " & r & ". ZF=" & GetReg("ZF") & " CF=" & GetReg("CF") & " SF=" & GetReg("SF")
End Sub

Public Sub BotonINC()
    Dim ax As Long
    Dim r As Long
    ax = GetReg("AX")
    r = ALU("INC", ax, 1)
    SetReg "AX", r
    Mensaje "INC: AX = " & ax & " + 1 = " & r & ". ZF=" & GetReg("ZF") & " CF=" & GetReg("CF") & " SF=" & GetReg("SF")
End Sub

Public Sub BotonDEC()
    Dim ax As Long
    Dim r As Long
    ax = GetReg("AX")
    r = ALU("DEC", ax, 1)
    SetReg "AX", r
    Mensaje "DEC: AX = " & ax & " - 1 = " & r & ". ZF=" & GetReg("ZF") & " CF=" & GetReg("CF") & " SF=" & GetReg("SF")
End Sub

Public Sub BotonAND()
    Dim ax As Long
    Dim bx As Long
    Dim r As Long
    ax = GetReg("AX")
    bx = GetReg("BX")
    r = ALU("AND", ax, bx)
    SetReg "AX", r
    Mensaje "AND: AX = " & ax & " AND " & bx & " = " & r & ". ZF=" & GetReg("ZF") & " CF=" & GetReg("CF") & " SF=" & GetReg("SF")
End Sub

Public Sub BotonOR()
    Dim ax As Long
    Dim bx As Long
    Dim r As Long
    ax = GetReg("AX")
    bx = GetReg("BX")
    r = ALU("OR", ax, bx)
    SetReg "AX", r
    Mensaje "OR: AX = " & ax & " OR " & bx & " = " & r & ". ZF=" & GetReg("ZF") & " CF=" & GetReg("CF") & " SF=" & GetReg("SF")
End Sub

Public Sub BotonXOR()
    Dim ax As Long
    Dim bx As Long
    Dim r As Long
    ax = GetReg("AX")
    bx = GetReg("BX")
    r = ALU("XOR", ax, bx)
    SetReg "AX", r
    Mensaje "XOR: AX = " & ax & " XOR " & bx & " = " & r & ". ZF=" & GetReg("ZF") & " CF=" & GetReg("CF") & " SF=" & GetReg("SF")
End Sub

Public Sub BotonNOT()
    Dim ax As Long
    Dim r As Long
    ax = GetReg("AX")
    r = ALU("NOT", ax, 0)
    SetReg "AX", r
    Mensaje "NOT: AX = NOT " & ax & " = " & r & ". ZF=" & GetReg("ZF") & " CF=" & GetReg("CF") & " SF=" & GetReg("SF")
End Sub

Public Sub BotonCMP()
    Dim ax As Long
    Dim bx As Long
    Dim r As Long
    ax = GetReg("AX")
    bx = GetReg("BX")
    r = ALU("CMP", ax, bx)
    Mensaje "CMP: compara " & ax & " y " & bx & " sin cambiar AX (sigue en " & GetReg("AX") & "). ZF=" & GetReg("ZF") & " CF=" & GetReg("CF") & " SF=" & GetReg("SF")
End Sub

Public Sub BotonResetCPU()
    SetReg "PC", 0
    SetReg "IR", 0
    SetReg "MAR", 0
    SetReg "MDR", 0
    SetReg "AX", 0
    SetReg "BX", 0
    HojaCPU.Range("E13").Value = 0
    HojaCPU.Range("G13").Value = 0
    HojaCPU.Range("I13").Value = 0
    PintarFlags
    Mensaje "Registros y banderas en cero."
End Sub

Public Sub BotonPruebaADD()
    SetReg "AX", 255
    SetReg "BX", 1
    BotonADD
    If GetReg("AX") = 0 And GetReg("ZF") = 1 And GetReg("CF") = 1 Then
        Mensaje "Prueba bien: 255 + 1 = 0. ZF=1 y CF=1. El resultado no cabe en 8 bits."
    Else
        Mensaje "Prueba mal: 255 + 1 deberia dejar AX=0, ZF=1, CF=1."
    End If
End Sub

Public Sub BotonPruebaSUB()
    SetReg "AX", 0
    SetReg "BX", 1
    BotonSUB
    If GetReg("AX") = 255 And GetReg("SF") = 1 And GetReg("CF") = 1 Then
        Mensaje "Prueba bien: 0 - 1 = 255. SF=1 y CF=1. Hubo prestamo y el bit 7 vale 1."
    Else
        Mensaje "Prueba mal: 0 - 1 deberia dejar AX=255, SF=1, CF=1."
    End If
End Sub

Public Sub BotonPruebaCMP()
    Dim axAntes As Long
    SetReg "AX", 7
    SetReg "BX", 7
    axAntes = GetReg("AX")
    BotonCMP
    If GetReg("AX") = axAntes And GetReg("ZF") = 1 Then
        Mensaje "Prueba bien: CMP AX, AX deja ZF=1 y AX sigue en " & GetReg("AX") & "."
    Else
        Mensaje "Prueba mal: CMP no debe cambiar AX y ZF debe ser 1."
    End If
End Sub
