' *****************************************************************************************
' Sub DibujarGanttEnFilasVacias
' Descripción:
'   Este procedimiento recorre las filas de una hoja de cálculo a partir de la fila 6 y,
'   si encuentra valores en la columna A, llama a otra función para llenar las celdas vacías
'   de un diagrama de Gantt con un valor específico.
' 
' Parámetros:
'   Ninguno
' 
' Notas:
'   - El rango del Gantt se asume que comienza en la celda K6 y se extiende hasta la última 
'     columna con datos en la fila 6.
'   - Esta función muestra un mensaje al finalizar el proceso.
' 
' *****************************************************************************************
Sub DibujarGanttSiFilaVacia()
    Dim ws As Worksheet
    Dim ultimaFilaA As Long
    Dim rangoGanttEncabezados As Range ' Rango de las fechas en la fila 4
    Dim primeraColumnaGantt As Long
    Dim ultimaColumnaGantt As Long
    Dim filaActual As Long
    Dim fechaInicioActividad As Date
    Dim fechaFinActividad As Date
    Dim columnaInicioGantt As Long
    Dim columnaFinGantt As Long
    Dim i As Long
    Dim celdaGanttFila As Range
    Dim estaFilaGanttVacia As Boolean

    ' Define la hoja de cálculo "Plan"
    Set ws = ThisWorkbook.Sheets("Plan") ' Asegúrate de que tu hoja se llama "Plan"

    ' Encuentra la última fila con datos en la columna A (a partir de la fila 6)
    ultimaFilaA = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Define el rango de los encabezados de fecha del Gantt (fila 4, desde la columna K hasta la última columna con fecha)
    Set rangoGanttEncabezados = ws.Range("K4", ws.Cells(4, Columns.Count).End(xlToLeft))
    primeraColumnaGantt = rangoGanttEncabezados.Column
    ultimaColumnaGantt = rangoGanttEncabezados.Column + rangoGanttEncabezados.Columns.Count - 1

    ' Itera a través de las filas desde la 6 hasta la última fila con valor en la columna A
    For filaActual = 6 To ultimaFilaA
        ' Verifica si la columna A de la fila actual tiene un valor
        If Not IsEmpty(ws.Cells(filaActual, "A").Value) Then
            ' Lee la fecha de inicio (columna E) y fin (columna D)
            On Error Resume Next
            fechaInicioActividad = DateValue(ws.Cells(filaActual, "D").Value)
            fechaFinActividad = DateValue(ws.Cells(filaActual, "E").Value)
            On Error GoTo 0

            ' Verifica si las fechas son válidas
            If IsDate(fechaInicioActividad) And IsDate(fechaFinActividad) Then
                ' Define el rango de la fila actual del Gantt (desde la primera hasta la última columna del encabezado)
                Dim rangoFilaGantt As Range
                Set rangoFilaGantt = ws.Range(ws.Cells(filaActual, primeraColumnaGantt), ws.Cells(filaActual, ultimaColumnaGantt))

                ' Verifica si la fila del Gantt está vacía
                estaFilaGanttVacia = True
                For Each celdaGanttFila In rangoFilaGantt
                    If Not IsEmpty(celdaGanttFila.Value) Then
                        estaFilaGanttVacia = False
                        Exit For
                    End If
                Next celdaGanttFila

                ' Si la fila del Gantt está vacía, procede a dibujar
                If estaFilaGanttVacia Then
                    ' Encuentra la columna de inicio en el Gantt
                    columnaInicioGantt = 0
                    For i = primeraColumnaGantt To ultimaColumnaGantt
                        If DateValue(ws.Cells(4, i).Value) = DateValue(fechaInicioActividad) Then
                            columnaInicioGantt = i
                            Exit For
                        End If
                    Next i

                    ' Encuentra la columna de fin en el Gantt
                    columnaFinGantt = 0
                    For i = primeraColumnaGantt To ultimaColumnaGantt
                        If DateValue(ws.Cells(4, i).Value) = DateValue(fechaFinActividad) Then
                            columnaFinGantt = i
                            Exit For
                        End If
                    Next i

                    ' Si se encontraron las columnas de inicio y fin, llena el rango con "x"
                    If columnaInicioGantt > 0 And columnaFinGantt > 0 Then
                        ' Asegurarse de que la columna de fin no sea anterior a la de inicio
                        If columnaFinGantt >= columnaInicioGantt Then
                            ws.Range(ws.Cells(filaActual, columnaInicioGantt), ws.Cells(filaActual, columnaFinGantt)).Value = "x"
                        Else
                            MsgBox "La fecha de fin es anterior a la fecha de inicio en la fila " & filaActual & ".", vbExclamation
                        End If
                    Else
                        If columnaInicioGantt = 0 Then
                            MsgBox "No se encontró la fecha de inicio en el encabezado del Gantt para la fila " & filaActual & ".", vbExclamation
                        End If
                        If columnaFinGantt = 0 Then
                            MsgBox "No se encontró la fecha de fin en el encabezado del Gantt para la fila " & filaActual & ".", vbExclamation
                        End If
                    End If
                End If
            Else
                ' Si las fechas no son válidas, podrías mostrar un mensaje o simplemente omitir la fila
                MsgBox "Fechas de inicio o fin no válidas en la fila " & filaActual & ".", vbExclamation
            End If
        End If
    Next filaActual

    MsgBox "Proceso de dibujo del Gantt completado.", vbInformation
End Sub
