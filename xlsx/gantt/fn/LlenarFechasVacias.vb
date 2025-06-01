' ***********************************************************************
' Función: LlenarFechasVacias
' Descripción:
'   Este procedimiento recorre todas las filas de una hoja de cálculo activa
'   y llena las celdas vacías de las columnas D y E con valores calculados
'   usando las funciones CalcDireccionMasIzquierda y CalcDireccionMasDerecha.
'
' Parámetros:
'   Ninguno.
'
' Requisitos:
'   - Las funciones CalcDireccionMasIzquierda y CalcDireccionMasDerecha deben estar
'     definidas en el mismo módulo o en otro módulo accesible.
'   - El procedimiento utiliza la hoja activa del libro. Si se requiere trabajar
'     con una hoja específica, se debe modificar el código para referirse a ella.
'
' Notas:
'   - Las columnas D y E son procesadas para llenar las celdas vacías.
'   - La función CalcDireccionMasIzquierda calcula la referencia de la celda no vacía
'     más a la izquierda de una fila, mientras que CalcDireccionMasDerecha realiza
'     lo mismo pero buscando desde la derecha.
'
' Resultados:
'   - Las celdas vacías en las columnas D y E se llenan con referencias de posiciones
'     calculadas.
'
' Mensajes:
'   - Al finalizar, muestra un mensaje indicando que el proceso ha sido completado.
' ***********************************************************************
Sub LlenarFechasVacias()

  Dim ws As Worksheet
  Dim ultimaFila As Long
  Dim i As Long
  Dim direccionIzquierda As String
  Dim direccionDerecha As String
  Dim valorIzquierdaFila4 As Variant
  Dim valorDerechaFila4 As Variant

  ' Especifica la hoja de cálculo que quieres procesar
  Set ws = ThisWorkbook.Sheets("Plan") ' Asegúrate de que tu hoja se llama "Plan"
  ' Encuentra la última fila con datos en alguna columna (ajusta la columna si es necesario)
  ultimaFila = ws.Cells(Rows.Count, "A").End(xlUp).Row

  ' Itera a través de cada fila desde la fila 1 hasta la última fila con datos
  For i = 6 To ultimaFila

    ' Para la columna D: ejecuta la función si la celda está vacía
    If Trim(ws.Cells(i, "D").Value) = "" Then
      direccionIzquierda = CalcDireccionMasIzquierda(i)
      If direccionIzquierda <> "" Then
        ' Extrae la letra de la columna de la dirección
        Dim columnaIzquierda As String
        columnaIzquierda = Split(direccionIzquierda, "$")(1)
        ' Obtiene el valor de la fila 4 de esa columna
        If ws.Rows(4).Columns(columnaIzquierda).Row = 4 Then
          valorIzquierdaFila4 = ws.Rows(4).Columns(columnaIzquierda).Value
          ws.Cells(i, "D").Value = valorIzquierdaFila4
        Else
          ws.Cells(i, "D").Value = "" ' O algún otro valor si no hay fila 4 válida
        End If
      Else
        ws.Cells(i, "D").Value = "" ' Si no se encontró dirección izquierda
      End If
    End If

    ' Para la columna E: ejecuta la función si la celda está vacía
    If Trim(ws.Cells(i, "E").Value) = "" Then
      direccionDerecha = CalcDireccionMasDerecha(i)
      If direccionDerecha <> "" Then
        ' Extrae la letra de la columna de la dirección
        Dim columnaDerecha As String
        columnaDerecha = Split(direccionDerecha, "$")(1)
        ' Obtiene el valor de la fila 4 de esa columna
        If ws.Rows(4).Columns(columnaDerecha).Row = 4 Then
          valorDerechaFila4 = ws.Rows(4).Columns(columnaDerecha).Value
          ws.Cells(i, "E").Value = valorDerechaFila4
        Else
          ws.Cells(i, "E").Value = "" ' O algún otro valor si no hay fila 4 válida
        End If
      Else
        ws.Cells(i, "E").Value = "" ' Si no se encontró dirección derecha
      End If
    End If

  Next i

  MsgBox "Proceso completado.", vbInformation

End Sub

' (Las funciones CalcDireccionMasIzquierda y CalcDireccionMasDerecha
'  deben estar definidas en el mismo módulo)

' Función CalcDireccionMasIzquierda
Function CalcDireccionMasIzquierda(fila As Long) As String

  Dim celda As Range
  Dim ultimaColumna As Long

  ultimaColumna = Cells(fila, Columns.Count).End(xlToLeft).Column

  For Each celda In Range(Cells(fila, "K"), Cells(fila, ultimaColumna))
    If Trim(celda.Value) <> "" Then
      CalcDireccionMasIzquierda = celda.Address
      Exit Function
    End If
  Next celda

  CalcDireccionMasIzquierda = ""

End Function

' Función CalcDireccionMasDerecha
Function CalcDireccionMasDerecha(fila As Long) As String

  Dim ultimaColumna As Long
  Dim i As Long

  ultimaColumna = Cells(fila, Columns.Count).End(xlToLeft).Column

  For i = ultimaColumna To Columns("K").Column Step -1
    If Trim(Cells(fila, i).Value) <> "" Then
      CalcDireccionMasDerecha = Cells(fila, i).Address
      Exit Function
    End If
  Next i

  CalcDireccionMasDerecha = ""

End Function
