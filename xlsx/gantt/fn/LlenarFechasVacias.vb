Sub LlenarFechasVacias()

  Dim ws As Worksheet
  Dim ultimaFila As Long
  Dim i As Long
  Dim resultadoIzquierda As String
  Dim resultadoDerecha As String

  ' Especifica la hoja de cálculo que quieres procesar
  Set ws = ThisWorkbook.ActiveSheet ' Puedes cambiar ActiveSheet por el nombre de tu hoja (ej: Worksheets("Hoja1"))

  ' Encuentra la última fila con datos en alguna columna (ajusta la columna si es necesario)
  ultimaFila = ws.Cells(Rows.Count, "A").End(xlUp).Row

  ' Itera a través de cada fila desde la fila 1 hasta la última fila con datos
  For i = 1 To ultimaFila

    ' Para la columna D: ejecuta la función si la celda está vacía
    If Trim(ws.Cells(i, "D").Value) = "" Then
      resultadoIzquierda = CalcDireccionMasIzquierda(i)
      ' Opcional: Puedes escribir el resultado en la columna D o hacer otra acción
      ws.Cells(i, "D").Value = resultadoIzquierda
    End If

    ' Para la columna E: ejecuta la función si la celda está vacía
    If Trim(ws.Cells(i, "E").Value) = "" Then
      resultadoDerecha = CalcDireccionMasDerecha(i)
      ' Opcional: Puedes escribir el resultado en la columna E o hacer otra acción
      ws.Cells(i, "E").Value = resultadoDerecha
    End If

  Next i

  MsgBox "Proceso completado.", vbInformation

End Sub

' (Las funciones CalcDireccionMasIzquierda y CalcDireccionMasDerecha
'  deben estar definidas en el mismo módulo o en un módulo diferente)

' Función CalcDireccionMasIzquierda (copia aquí si no la tienes en el módulo)
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

' Función CalcDireccionMasDerecha (copia aquí si no la tienes en el módulo)
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
