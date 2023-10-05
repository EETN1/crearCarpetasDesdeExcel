# crearCarpetasDesdeExcel
Utilizando los macros creamos un script para crear carpetas en base a los datos de una columna


    
    Sub CrearCarpetasDesdeColumna()
    Dim Rng As Range
    Dim Cel As Range
    Dim CarpetaBase As String
    Dim NuevaCarpeta As String
    
    On Error Resume Next ' Ignora los errores y continúa
    
    ' Especifica la columna en la que se encuentran los nombres de las carpetas
     Set Rng = ThisWorkbook.Sheets("Asistencias").Range("B4:B" & ThisWorkbook.Sheets("Asistencias").Cells(Rows.Count, 1).End(xlUp).Row)
    
    ' Especifica la carpeta base donde deseas crear las subcarpetas
    'CarpetaBase = "C:\prueba" ' Cambia la ruta a la carpeta base que desees
      ' Solicitar al usuario la ruta de destino
    CarpetaBase = InputBox("Ingrese la ruta de destino para las carpetas:", "Ruta de Destino")
    
    ' Recorre cada celda en la columna y crea una carpeta con el nombre de la celda
    For Each Cel In Rng
        If Len(Cel.Value) > 0 Then
            NuevaCarpeta = CarpetaBase & "\" & Cel.Value
            MkDir NuevaCarpeta ' Intenta crear la carpeta
            
            ' Manejo de errores
            If Err.Number <> 0 Then
                MsgBox "Error al crear la carpeta: " & NuevaCarpeta, vbExclamation, "Error"
                Err.Clear ' Limpia el error
            End If
        End If
    Next Cel
    
    On Error GoTo 0 ' Restablece el manejo de errores
    End Sub
