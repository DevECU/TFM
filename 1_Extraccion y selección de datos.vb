Public Sub main()

' Módulo    : Extracción y selección de datos
' Autor     : Bryan Mayorga
' Fecha     : 2022/06/02
' Propósito : Obtiene datos de medidores inteligentes y selecciona variable de energía

'Desactivo parpadeo
Application.ScreenUpdating = False 'Desactiva parpadeo
Application.EnableCancelKey = xlDisabled

Dim numero_archivo As Integer
Dim carpeta, archivos As String
Dim contador As Integer
Dim archivoconsultado As String
Dim nombre_archivo As String
Dim archivoabrir As Excel.Workbook

Dim ultima_fila As Long

Dim i As Long
Dim j, z, x, y, k, celdas_con_datos
Dim MiMatriz(1 To 1000000, 1 To 2) As String

Dim Fecha_inicial

'''''''''''Extracción de nombre de archivos''''''''''''
'Borro datos previos
Range("A1:A1000000").Select
Selection.ClearContents
Range("A1").Select

carpeta = InputBox("Ingrese la ruta del archivo importar")

If carpeta = "" Then
    Exit Sub
ElseIf Right(carpeta, 1) <> "\" Then
    carpeta = carpeta & "\"
End If

contador = 1
archivos = Dir(carpeta)

Do While Len(archivos) > 0
    ActiveSheet.Cells(contador, 1).Value = archivos
    archivos = Dir()
    contador = contador + 1
Loop

nuevo_archivo = 0
z = 1
Sheets("Data").Select
ActiveSheet.Range("A1:ES300000").Select
Selection.ClearContents
ActiveSheet.Cells(1, 1).Select
Sheets("Archivos").Select

'Abro uno por uno los archivos
For numero_archivo = 1 To contador   'contador define el número de archivos

    nombre_archivo = Cells(numero_archivo, 1)
    archivoconsultado = carpeta & "\" & nombre_archivo
    If Len(archivoconsultado) > 0 Then
        Set archivoabrir = Workbooks.Open(archivoconsultado)
        ultima_fila = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
        
        'Copio datos
        For i = 1 To ultima_fila
            For j = 1 To 2 'Número de columnas a ser copiadas
                MiMatriz(i, j) = ActiveSheet.Cells(3 + i, j).Value 'El 4 le da el offset
            Next j
        Next i
        Workbooks.Open(archivoconsultado).Close
                
        Sheets("Data").Select

        'Pego datos      
        Fecha_inicial = CDate("1/1/2021  0:10:00")
      
        If MiMatriz(1, 1) = Fecha_inicial Then
                ActiveSheet.Cells(1, z + 1).Value = nombre_archivo
                For i = 1 To ultima_fila - 4
                    j = 1
                        If MiMatriz(i, j) <> "" Then
                            ActiveSheet.Cells(i + 1, j).Value = CDate(MiMatriz(i, j))
                        End If
                        If MiMatriz(i, j + 1) <> "" Then
                            ActiveSheet.Cells(i + 1, j + z).Value = CDbl(MiMatriz(i, j + 1))
                        End If
                Next i
                z = z + 1
          End If
    
    Else
        responde = MsgBox("El archivo no existe en esta carpeta")
        If responde = vbYes Then
            Miarchivo = Application.geopenfilename
            Workbooks.Open Miarchivo
        End If
    End If
Next numero_archivo
End Sub
