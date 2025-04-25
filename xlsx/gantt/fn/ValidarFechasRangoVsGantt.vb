Sub ValidarFechasVsGanttThenRemarcar()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim rangoGantt As Range
    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim primeraColumnaGantt As Long
    Dim ultimaColumnaGantt As Long
    Dim celdaInicioGantt As Range
    Dim celdaFinGantt As Range
    Dim i As Long
    Dim j As Long
    Dim encontradoInicio As Boolean
    Dim encontradoFin As Boolean
    Const ESTILO_INCORRECTO As String = "Incorrecto"  ' Define el nombre del estilo


    ' Define la hoja de cálculo donde está tu Gantt
    Set ws = ThisWorkbook.Sheets("Plan") ' Puedes cambiarlo por Sheets("NombreDeTuHoja")

    ' Encuentra la última fila con datos en la columna E (Fecha Inicial)
    ultimaFila = ws.Cells(Rows.Count, "E").End(xlUp).Row

    ' Define el rango donde se dibuja el Gantt
    Set rangoGantt = ws.Range("K6", ws.Cells(ultimaFila, ws.Columns.Count).End(xlToLeft))

    ' Obtiene la primera y última columna del rango Gantt
    primeraColumnaGantt = rangoGantt.Column
    ultimaColumnaGantt = rangoGantt.Columns.Count + primeraColumnaGantt - 1

    ' Recorre cada fila de actividades
    For i = 6 To ultimaFila
        ' Obtiene las fechas de inicio y fin
        On Error Resume Next ' En caso de que las celdas de fecha estén vacías o no sean fechas válidas
        fechaInicio = DateValue(ws.Cells(i, "D").Value)
        fechaFin = DateValue(ws.Cells(i, "E").Value)
        On Error GoTo 0

        ' Limpia el color de fondo de las celdas de fecha
        ws.Cells(i, "D").Interior.ColorIndex = xlNone
        ws.Cells(i, "E").Interior.ColorIndex = xlNone

        ' Busca la primera celda no vacía en la fila del Gantt
        Set celdaInicioGantt = Nothing
        encontradoInicio = False
        For j = primeraColumnaGantt To ultimaColumnaGantt
            If Not IsEmpty(ws.Cells(i, j).Value) Then
                Set celdaInicioGantt = ws.Cells(i, j)
                encontradoInicio = True
                Exit For
            End If
        Next j

        ' Busca la última celda no vacía en la fila del Gantt
        Set celdaFinGantt = Nothing
        encontradoFin = False
        If encontradoInicio Then
            For j = ultimaColumnaGantt To primeraColumnaGantt Step -1
                If Not IsEmpty(ws.Cells(i, j).Value) Then
                    Set celdaFinGantt = ws.Cells(i, j)
                    encontradoFin = True
                    Exit For
                End If
            Next j
        End If

        ' Realiza la validación
        If encontradoInicio And encontradoFin Then
            If DateValue(fechaInicio) <> DateValue(ws.Cells(4, celdaInicioGantt.Column).Value) Then
                ws.Cells(i, "E").Style = ESTILO_INCORRECTO ' Colorea la fecha de inicio
            End If
            If DateValue(fechaFin) <> DateValue(ws.Cells(4, celdaFinGantt.Column).Value) Then
                ws.Cells(i, "D").Style = ESTILO_INCORRECTO ' Colorea la fecha de fin
            End If
        ElseIf (encontradoInicio And Not encontradoFin) Or (Not encontradoInicio And encontradoFin) Then
            ' Si solo hay un extremo marcado en el Gantt y hay fechas, también marcar
            If IsDate(ws.Cells(i, "E").Value) Then ws.Cells(i, "E").Style = ESTILO_INCORRECTO
            If IsDate(ws.Cells(i, "D").Value) Then ws.Cells(i, "D").Style = ESTILO_INCORRECTO
        Else ' Not encontradoInicio And Not encontradoFin
            ' Si no hay marcas en el Gantt, puedes decidir si las fechas son válidas o no.
            ' Por ahora, no se colorea nada en este caso. Puedes agregar lógica si deseas marcar
            ' las fechas en rojo incluso si no hay marcas en el Gantt.
        End If
    Next i

    MsgBox "Validación del Gantt completada.", vbInformation
End Sub
