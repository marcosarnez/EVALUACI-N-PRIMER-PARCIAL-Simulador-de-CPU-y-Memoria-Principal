# Simulador de CPU von Neumann de 8 bits

Proyecto del primer parcial de Arquitectura de Computadoras (SIS-131), Universidad Católica Boliviana San Pablo. El simulador corre en Excel con macros VBA (`.xlsm`). Modela una CPU de 8 bits y una RAM de 256 bytes. Cada **STEP** ejecuta una sola fase del ciclo: Fetch, Decode, Execute o Store.

## Cómo abrirlo

1. Abre `simulador-cpu.xlsm`.
2. Si Excel avisa, habilita las macros. Si el archivo llegó bloqueado: clic derecho, Propiedades, Desbloquear.
3. Trabaja en la hoja **CPU**. La hoja **RAM** guarda el byte de cada casilla (fila 1 = dirección 0). No la edites a mano. La hoja **Inspección** lista las 256 casillas en hexadecimal, binario y decimal.

## Arquitectura

```mermaid
flowchart LR
    PC["PC"] --> MAR["MAR"]
    MAR --> RAM["RAM 256 x 8 bits 00h-FFh"]
    RAM --> MDR["MDR"]
    MDR --> IR["IR"]
    IR --> UC["Unidad de control"]
    UC --> ALU["ALU"]
    AX["AX"] --> ALU
    BX["BX"] --> ALU
    ALU --> FLAGS["ZF CF SF"]
    ALU --> AX
    ALU --> BX
    MDR --> RAM
```

Una sola memoria guarda el programa (`00h`–`7Fh`) y los datos (`80h`–`FFh`). El CPU no escribe esa memoria por su cuenta: solo `ReadRAM` y `WriteRAM`.

## Manual

1. Pulsa **LOAD FIBO** (programa obligatorio) o **LOAD 5+3** (ejemplo `a = 5`, `b = 3`, `c = a + b`).
2. **STEP** avanza una fase. Cuatro clics completan una instrucción. La fase encendida y la columna **Estado** marcan qué está trabajando. La celda amarilla del mapa es la casilla activa.
3. El mnemónico (por ejemplo `LOAD AX,[0x80]`) aparece una sola vez, en **Celda seleccionada**.
4. **RUN** sigue solo. **PAUSE** congela el reloj. El retardo está en **Delay (ms)** (250 va bien en la defensa).
5. **RESET** pone PC, IR, MAR, MDR, AX, BX y las banderas en cero, apaga **ACTIVO** y vacía el log. No borra la RAM.
6. **HLT** detiene STEP y RUN. Para seguir, carga de nuevo o pulsa RESET.
7. El log muestra lo más reciente arriba, con la forma `[Paso NN] FASE: detalle`. Hay una sola línea de estado.

## ISA

El byte de registro vale `00` para AX y `01` para BX.

| Mnemónico | Opcode | Bytes | Operandos | Descripción | Flags |
|-----------|--------|-------|-----------|-------------|-------|
| `HLT` | `00` | 1 | — | Detiene el reloj | — |
| `MOV r, imm` | `01` | 3 | reg, imm8 | Copia el inmediato al registro | — |
| `MOV r, r` | `02` | 3 | reg, reg | Copia un registro a otro | — |
| `LOAD r, [dir]` | `03` | 3 | reg, dir8 | Lee la RAM hacia el registro | — |
| `STORE [dir], r` | `04` | 3 | dir8, reg | Escribe el registro en la RAM | — |
| `ADD r, imm` | `10` | 3 | reg, imm8 | Suma un inmediato | ZF, CF, SF |
| `ADD r, r` | `11` | 3 | reg, reg | Suma un registro | ZF, CF, SF |
| `SUB r, imm` | `12` | 3 | reg, imm8 | Resta un inmediato | ZF, CF, SF |
| `SUB r, r` | `13` | 3 | reg, reg | Resta un registro | ZF, CF, SF |
| `INC r` | `14` | 2 | reg | Incrementa | ZF, CF, SF |
| `DEC r` | `15` | 2 | reg | Decrementa | ZF, CF, SF |
| `CMP r, imm` | `16` | 3 | reg, imm8 | Resta sin guardar el resultado | ZF, CF, SF |
| `CMP r, r` | `17` | 3 | reg, reg | Resta sin guardar el resultado | ZF, CF, SF |
| `AND r, r` | `18` | 3 | reg, reg | AND bit a bit | ZF, SF (CF = 0) |
| `OR r, r` | `19` | 3 | reg, reg | OR bit a bit | ZF, SF (CF = 0) |
| `XOR r, r` | `1A` | 3 | reg, reg | XOR bit a bit | ZF, SF (CF = 0) |
| `NOT r` | `1B` | 2 | reg | NOT bit a bit | ZF, SF (CF = 0) |
| `JMP dir` | `20` | 2 | dir8 | PC = dir | — |
| `JZ dir` | `21` | 2 | dir8 | Salta si ZF = 1 | — |
| `JNZ dir` | `22` | 2 | dir8 | Salta si ZF = 0 | — |
| `JC dir` | `23` | 2 | dir8 | Salta si CF = 1 | — |

**ZF** = 1 si el resultado es 0. **CF** = 1 si hubo acarreo o préstamo fuera de 8 bits. **SF** = bit 7 del resultado.

Pruebas de banderas, en el panel de la ALU: `255 + 1` deja AX = 0, ZF = 1, CF = 1. `0 - 1` deja AX = 255, CF = 1, SF = 1. `CMP AX, AX` deja ZF = 1 y no cambia AX.

## Ciclo de instrucción

1. **Fetch.** `PC → MAR`, `ReadRAM → MDR → IR`, el PC avanza el tamaño de la instrucción.
2. **Decode.** La unidad de control lee el opcode del IR y prepara los operandos. Ahí aparece el mnemónico.
3. **Execute.** La ALU opera, o se decide un salto, y se actualizan las banderas.
4. **Store.** El resultado se escribe en AX, en BX o en la RAM. `CMP` no escribe. `HLT` para el reloj.

## Traza del Fibonacci

Programa en `00h`–`13h`. Datos: `80h` = a, `81h` = b. Arranque: a = 0, b = 1. AX = 0, BX = 0, ZF = 0, CF = 0, SF = 0.

```
00  LOAD AX,[0x80]
03  LOAD BX,[0x81]
06  ADD  AX,BX
09  JC   0x13
0B  STORE [0x80],BX
0E  STORE [0x81],AX
11  JMP  0x00
13  HLT
```

Bytes: `03 00 80 03 01 81 11 00 01 23 13 04 01 80 04 00 81 20 00 00`.

Primera vuelta, al terminar cada instrucción (después de Store):

| Instrucción | PC | IR | AX | BX | ZF | CF | SF | Memoria |
|-------------|----|----|----|----|----|----|----|---------|
| inicio | 00 | — | 0 | 0 | 0 | 0 | 0 | 80h=0, 81h=1 |
| `LOAD AX,[0x80]` | 03 | 03 | 0 | 0 | 0 | 0 | 0 | igual |
| `LOAD BX,[0x81]` | 06 | 03 | 0 | 1 | 0 | 0 | 0 | igual |
| `ADD AX,BX` | 09 | 11 | 1 | 1 | 0 | 0 | 0 | igual |
| `JC 0x13` | 0B | 23 | 1 | 1 | 0 | 0 | 0 | CF=0, no salta |
| `STORE [0x80],BX` | 0E | 04 | 1 | 1 | 0 | 0 | 0 | 80h=1 |
| `STORE [0x81],AX` | 11 | 04 | 1 | 1 | 0 | 0 | 0 | 81h=1 |
| `JMP 0x00` | 00 | 20 | 1 | 1 | 0 | 0 | 0 | vuelve al inicio |

La serie sigue hasta guardar 144 y 233. La suma siguiente es 144 + 233 = 377. No cabe en 8 bits: AX queda 121, CF = 1, SF = 0, ZF = 0. `JC 0x13` sí salta y `HLT` detiene el reloj. Las casillas quedan `80h` = 144 y `81h` = 233.

| Momento | PC | IR | AX | BX | ZF | CF | SF |
|---------|----|----|----|----|----|----|----|
| `ADD` de 144+233, tras Store | 09 | 11 | 121 | 233 | 0 | 1 | 0 |
| `JC 0x13` tomado | 13 | 23 | 121 | 233 | 0 | 1 | 0 |
| `HLT` | 14 | 00 | 121 | 233 | 0 | 1 | 0 |

## Estructura

Archivos del simulador, junto al libro:

- `simulador-cpu.xlsm` — libro con las hojas CPU, RAM e Inspección
- `CPU.bas` — registros, ALU y banderas
- `UI.bas` — la única rutina que escribe en la hoja CPU
- `Ciclo.bas` — Fetch, Decode, Execute, Store, LOAD, RUN, PAUSE, RESET
- `MemoriaRAM.bas` — `ReadRAM` y `WriteRAM`
- `ThisWorkbook.cls` — protección de las hojas al abrir
- `README.md` — este manual

## Código para la defensa

| Pregunta | Dónde | Qué decir |
|----------|--------|-----------|
| Cómo se lee la RAM | `MemoriaRAM.bas`, `ReadRAM` | Pone la dirección en MAR, el byte en MDR y lo devuelve. |
| Cómo se escribe | `WriteRAM` | Es la única rutina que cambia la hoja RAM. |
| Cómo avanza el reloj | `Ciclo.bas`, `BotonSTEP` | Según la fase llama Fetch, Decode, Execute o Store. |
| Dónde se suman 8 bits | `CPU.bas`, `ALU` | Si la suma pasa de 255, CF = 1 y se guarda el byte bajo. |
| Dónde salta el Fibonacci | `HacerExecute`, caso `JC` | Si CF = 1, PC = destino. |
