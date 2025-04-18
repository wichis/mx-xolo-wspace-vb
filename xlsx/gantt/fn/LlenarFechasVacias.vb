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
  Dim resultadoIzquierda As String
  Dim resultadoDerecha As String
  Dim celdaIzq As Range
  Dim celdaDer As Range

  ' Especifica la hoja de cálculo que quieres procesar
  Set ws = ThisWorkbook.ActiveSheet ' Puedes cambiar ActiveSheet por el nombre de tu hoja (ej: Worksheets("Hoja1"))

  ' Encuentra la última fila con datos en alguna columna (ajusta la columna si es necesario)
  ultimaFila = ws.Cells(Rows.Count, "A").End(xlUp).Row

  ' Itera a través de cada fila desde la fila 1 hasta la última fila con datos
  For i = 6 To ultimaFila

    ' Para la columna D: ejecuta la función si la celda está vacía
    If Trim(ws.Cells(i, "D").Value) = "" Then
      resultadoIzquierda = CalcDireccionMasIzquierda(i)

      Set celdaIzq = Range(resultadoIzquierda)
      
      ' Escribe el resultado en la columna D
      ws.Cells(i, "D").Value = Cells(4, celdaIzq.Column).Value
    End If

    ' Para la columna E: ejecuta la función si la celda está vacía
    If Trim(ws.Cells(i, "E").Value) = "" Then
      resultadoDerecha = CalcDireccionMasDerecha(i)
      
      Set celdaDer = Range(resultadoDerecha)
      
      ' Escribe el resultado en la columna E
      ws.Cells(i, "E").Value = Cells(4, celdaDer.Column).Value
    End If

  Next i

  ' Muestra un mensaje al completar el proceso
  MsgBox "Proceso completado.", vbInformation

End Sub

' ***********************************************************************
' Función: CalcDireccionMasIzquierda
' Descripción:
'   Encuentra la dirección de la celda no vacía más a la izquierda de una fila
'   dentro de un rango específico, empezando desde la columna K.
'
' Parámetros:
'   - fila (Long): Número de la fila a analizar.
'
' Retorno:
'   - Dirección de la celda no vacía más a la izquierda (String).
'   - Retorna una cadena vacía si no se encuentra ninguna celda no vacía.
' ***********************************************************************
Function CalcDireccionMasIzquierda(fila As Long) As String

  Dim celda As Range
  Dim ultimaColumna As Long

  ' Encuentra la última columna no vacía de la fila
  ultimaColumna = Cells(fila, Columns.Count).End(xlToLeft).Column

  ' Recorre las celdas desde la columna K hasta la última columna no vacía
  For Each celda In Range(Cells(fila, "K"), Cells(fila, ultimaColumna))
    If Trim(celda.Value) <> "" Then
      CalcDireccionMasIzquierda = celda.Address
      Exit Function
    End If
  Next celda

  ' Si no se encuentra ninguna celda no vacía, retorna una cadena vacía
  CalcDireccionMasIzquierda = ""

End Function

' ***********************************************************************
' Función: CalcDireccionMasDerecha
' Descripción:
'   Encuentra la dirección de la celda no vacía más a la derecha de una fila
'   dentro de un rango específico, empezando desde la última columna no vacía.
'
' Parámetros:
'   - fila (Long): Número de la fila a analizar.
'
' Retorno:
'   - Dirección de la celda no vacía más a la derecha (String).
'   - Retorna una cadena vacía si no se encuentra ninguna celda no vacía.
' ***********************************************************************
Function CalcDireccionMasDerecha(fila As Long) As String

  Dim ultimaColumna As Long
  Dim i As Long

  ' Encuentra la última columna no vacía de la fila
  ultimaColumna = Cells(fila, Columns.Count).End(xlToLeft).Column

  ' Recorre las celdas desde la última columna no vacía hacia la columna K
  For i = ultimaColumna To Columns("K").Column Step -1
    If Trim(Cells(fila, i).Value) <> "" Then
      CalcDireccionMasDerecha = Cells(fila, i).Address
      Exit Function
    End If
  Next i

  ' Si no se encuentra ninguna celda no vacía, retorna una cadena vacía
  CalcDireccionMasDerecha = ""

End Function
