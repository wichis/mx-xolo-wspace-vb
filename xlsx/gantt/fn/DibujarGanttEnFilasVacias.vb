Sub DibujarGanttEnFilasVacias()
    Dim ws As Worksheet
    Dim ultimaFilaA As Long
    Dim rangoGantt As Range
    Dim filaActual As Long

    ' Define la hoja de cálculo
    Set ws = ThisWorkbook.ActiveSheet ' Puedes cambiarlo por Sheets("NombreDeTuHoja")

    ' Encuentra la última fila con datos en la columna A
    ultimaFilaA = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Define el rango donde se dibuja el Gantt (ajusta si es necesario)
    ' Asumimos que el rango Gantt comienza en K6 y se extiende hasta la última columna con datos en cualquier fila
    Dim ultimaColumnaGantt As Long
    ultimaColumnaGantt = ws.Cells(6, Columns.Count).End(xlToLeft).Column ' Encuentra la última columna con datos en la fila 6 (puedes ajustarlo si es necesario)
    Set rangoGantt = ws.Range("K6", ws.Cells(ultimaFilaA, ultimaColumnaGantt))

    ' Itera a través de las filas con valores en la columna A (a partir de la fila 6)
    For filaActual = 6 To ultimaFilaA
        ' Verifica si hay un valor en la columna A de la fila actual
        If Not IsEmpty(ws.Cells(filaActual, "A").Value) Then
            ' Llama a la función para llenar el Gantt si la fila está vacía
            Call LlenarGanttSiVacio(filaActual, rangoGantt, "E", "D", "x")
        End If
    Next filaActual

    MsgBox "Proceso de llenado del Gantt completado para filas con valores en la columna A.", vbInformation
End Sub

Sub LlenarGanttSiVacio(fila As Long, rangoGantt As Range, columnaInicioFechas As String, columnaFinFechas As String, valorLlenado As String)
    Dim ws As Worksheet
    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim fechaInicioGantt As Date
    Dim columnaInicioGantt As Long
    Dim columnaFinGantt As Long
    Dim i As Long
    Dim celdaGantt As Range
    Dim estaVacio As Boolean

    ' Define la hoja de cálculo
    Set ws = rangoGantt.Parent

    ' Obtiene las fechas de inicio y fin de la fila especificada
    On Error Resume Next
    fechaInicio = DateValue(ws.Cells(fila, columnaInicioFechas).Value)
    fechaFin = DateValue(ws.Cells(fila, columnaFinFechas).Value)
    On Error GoTo 0

    ' Si las fechas no son válidas, sale de la función
    If Not IsDate(fechaInicio) Or Not IsDate(fechaFin) Then
        ' Puedes optar por no mostrar un mensaje aquí si se procesan muchas filas automáticamente
        ' MsgBox "Fechas de inicio o fin no válidas en la fila " & fila, vbExclamation
        Exit Sub
    End If

    ' Obtiene la fecha de inicio del calendario del Gantt (fila 4, primera columna del rango Gantt)
    fechaInicioGantt = ws.Cells(4, rangoGantt.Column).Value

    ' Calcula las columnas de inicio y fin en el Gantt
    columnaInicioGantt = rangoGantt.Column + DateDiff("d", fechaInicioGantt, fechaInicio)
    columnaFinGantt = rangoGantt.Column + DateDiff("d", fechaInicioGantt, fechaFin)

    ' Verifica si la fila del Gantt está vacía
    estaVacio = True
    For Each celdaGantt In ws.Rows(fila).Range(rangoGantt.Address)
        If Not IsEmpty(celdaGantt.Value) Then
            estaVacio = False
            Exit For
        End If
    Next celdaGantt

    ' Llena la fila del Gantt si está vacía y las columnas calculadas están dentro del rango
    If estaVacio Then
        If columnaInicioGantt >= rangoGantt.Column And columnaFinGantt <= rangoGantt.Column + rangoGantt.Columns.Count - 1 Then
            For i = columnaInicioGantt To columnaFinGantt
                ws.Cells(fila, i).Value = valorLlenado
            Next i
        Else
            ' Puedes optar por no mostrar un mensaje aquí si se procesan muchas filas automáticamente
            ' MsgBox "Las fechas de la fila " & fila & " están fuera del rango del Gantt.", vbExclamation
        End If
    Else
        ' Puedes optar por no mostrar un mensaje aquí si se procesan muchas filas automáticamente
        ' MsgBox "La fila " & fila & " del Gantt ya contiene valores.", vbInformation
    End If

End Sub
