Sub carga_mediador()
   
' para el destino donde se copiara
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set LibroDestino = Workbooks(ThisWorkbook.Name)
    Set HojaDestino = LibroDestino.Worksheets("Hoja1")
    'ActiveWindow.Visible = False

    'carpeta_verificable = ActiveWorkbook.Path & "\Reportes\"
    carpeta_verificable = "D:\AD\BUN\cuadre\Mediador\"
    archivo = Dir(carpeta_verificable & "\*.x*")
    mover_fila = 1
    Do While Len(archivo) - 2 > 0
        ruta = carpeta_verificable & archivo

         ' para navegar en los archivos
        
        
            Set LibroOrigen = Workbooks.Open(ruta)
            ActiveWindow.Visible = False
        On Error GoTo abc
            Set HojaOrigen = LibroOrigen.Worksheets("ReporteDetalle")
        
        ' Call Tama?oMatriz
            varia = 10
            filas = 0
        
                Do While HojaOrigen.Cells(varia, 2) <> ""
                            varia = varia + 1
                            filas = filas + 1
                Loop
        ' Dimensionar la Matriz para que no quede rastros de los datos anteriores
                    ReDim matriz(filas, 7)
        ' Selecciona datos standares
                    agencia = HojaOrigen.Range("b6").Value
                    cliente = HojaOrigen.Range("b7").Value
        ' carga datos a la matriz OJO QUE SOLO SON DATOS SELECCIONADOS
                        For i = 1 To filas
                            For j = 1 To 7
                                If j = 1 Then
                                    matriz(i, j) = "Mediador"
                                ElseIf j = 2 Then
                                    matriz(i, j) = agencia
                                ElseIf j = 5 Then
                                    matriz(i, j) = HojaOrigen.Cells(i + 9, j + 2).Value
                                ElseIf j = 6 Then
                                    matriz(i, j) = HojaOrigen.Cells(i + 9, j + 2).Value
                                ElseIf j = 7 Then
                                    matriz(i, j) = cliente
                                Else
                                    matriz(i, j) = HojaOrigen.Cells(i + 9, j - 1).Value
                                End If
                            Next j
                        Next i
        ' Seleccionamos la hoja de destino y vaciamos la matriz
        
        
                HojaDestino.Activate
        
                        For i = 1 To filas
                            For j = 1 To 6
                                    HojaDestino.Cells(i + mover_fila, j).Value = matriz(i, j)
                            Next j
                        Next i
        ' Para que vaya llenando hacia abajo NO SE PUEDO UTILIZAR ACTIVECELL.OFFSET
                        mover_fila = mover_fila + i - 1
        ' Selecciona el siguiente archivo
abc:
            
            Workbooks(LibroOrigen.Name).Close Savechanges:=False
            archivo = Dir()
Loop
    
end sub
sub carga_unibanca
Dim i, j, fn As Integer
    Dim cadena, ruta As String
    Dim arr
    Dim matriz() As Variant
    
' para no mostrar acciones de excel
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
' para activar los libros virtuales
    Set HojaDestino = Workbooks(ThisWorkbook.Name).Worksheets("Hoja1")
    HojaDestino.Activate
    Worksheets.Add.Name = "Carga"

' Declarar los libros
    Set LibroCarga = Workbooks(ThisWorkbook.Name)
    Set HojaCarga = LibroCarga.Worksheets("Carga")


    HojaCarga.Activate

    fn = FreeFile
' Por el momento colocar la ruta del archivo
    ruta = "D:\AD\BUN\cuadre\UNIBANCA\unibanca.txt"
' Cargar el archivo a la variable fn
    Open ruta For Input As #fn

    Set LibroOrigen = ActiveWorkbook
'Recorre cada fila cargando al array y separando por ";"
    Do While Not EOF(fn)
        Line Input #fn, cadena
        i = i + 1
        arr = Split(cadena, ";")
            For j = 0 To UBound(arr)
                Sheets("Carga").Cells(i, j + 1).Value = arr(j)
            Next j
    Loop
' Dimensiones a la a la matriz
        varia = 1
        filas = 0
        columnas = 0

            Do While Worksheets("Carga").Cells(varia, 1) <> ""
                varia = varia + 1
                filas = filas + 1
            Loop
        varia = 1
            Do While Worksheets("Carga").Cells(1, varia) <> ""
                varia = varia + 1
                columnas = columnas + 1
            Loop
    ReDim matriz(filas, columnas)
' Carga de los datos a la matriz

        For i = 1 To filas
            For j = 1 To columnas
                matriz(i, j) = Worksheets("Carga").Cells(i, j).Value
            Next j
        Next i
Sheets("Carga").Delete
' Activamos la hoja de Destino

    HojaDestino.Activate
    Range("A1").Select

                Do While ActiveCell.Value <> ""
                    ActiveCell.Offset(1, 0).Select
                Loop

            For i = 1 To filas
                For j = 1 To 6
                    If j = 1 Then
                        ActiveCell.Value = "Unibanca"
                    'matriz(i, j) = Worksheets("Carga").Cells(i, j).Value
                    End If
                    If j = 2 Then
                        ActiveCell.Offset(0, 1).Value = matriz(i, 8) & matriz(i, 9) & matriz(i, 2) & matriz(i, 3)
                    End If
                    If j = 3 Then
                        ActiveCell.Offset(0, 2).Value = matriz(i, 19)
                    End If
                    If j = 4 Then
                        ActiveCell.Offset(0, 3).Value = matriz(i, 5)
                    End If
                    If j = 5 Then
                        ActiveCell.Offset(0, 4).Value = matriz(i, 21)
                    End If
                    If j = 6 Then
                        ActiveCell.Offset(0, 6).Value = matriz(i, 20)
                    End If
                Next j
                ActiveCell.Offset(1, 0).Select
            Next i
            
            Worksheets("Hoja2").Activate
 ActiveWindow.Visible = True
End Sub

Sub ImportarArchivodeTextoMediador()
    Dim i, j, fn As Integer
    Dim cadena, ruta As String
    Dim arr
    Dim matriz() As Variant
' para no mostrar acciones de excel
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
' para activar los libros virtuales
    Worksheets.Add.Name = "Carga"

' Declarar los libros
    Set LibroCarga = Workbooks(ThisWorkbook.Name)
    Set HojaCarga = LibroCarga.Worksheets("Carga")


    HojaCarga.Activate

    fn = FreeFile
' Por el momento colocar la ruta del archivo
    ruta = "D:\AD\BUN\cuadre\Mediador\mediador.txt"
' Cargar el archivo a la variable fn
    Open ruta For Input As #fn

    Set LibroOrigen = ActiveWorkbook
'Recorre cada fila cargando al array y separando por "|"
    Do While Not EOF(fn)
        Line Input #fn, cadena
        i = i + 1
        arr = Split(cadena, "|")
            For j = 0 To UBound(arr)
                Sheets("Carga").Cells(i, j + 1).Value = arr(j)
            Next j
    Loop
' Eliminando la primera fila
    If ActiveCell.Value = "" Then
        Rows(1).EntireRow.Delete
    End If

 ' Para la limpieza de la base
    varia = 1
            Do While Worksheets("Carga").Cells(varia, 1) <> ""
                    If IsNumeric(Cells(varia, 4).Value) Then
                        If Cells(varia, 4).Value = "" Then
                            Rows(varia).EntireRow.Delete
                            varia = varia - 1
                        End If
                    Else
                        Rows(varia).EntireRow.Delete
                        varia = varia - 1
                    End If
                    varia = varia + 1
            Loop

' Dimensiones a la a la matriz

                filas = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
                columnas = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    ReDim matriz(filas, columnas)
' Carga de los datos a la matriz
        For i = 1 To filas
            For j = 1 To columnas
                matriz(i, j) = Worksheets("Carga").Cells(i, j).Value
            Next j
        Next i
Sheets("Carga").Delete
' Imprimir la Matriz en la Base (Hoja1) // COLOCAR CABEZALES
    Set HojaDestino = Workbooks(ThisWorkbook.Name).Worksheets("Hoja1")
    HojaDestino.Activate
    Range("A2").Select
               
                    For i = 1 To filas
                        For j = 1 To 6
                            If j = 1 Then
                                Cells(i + 1, j).Value = "Mediador"
                            End If
                            If j = 2 Then
                                Cells(i + 1, j).Value = ""
                            End If
                            If j = 3 Then
                                Cells(i + 1, j).Value = matriz(i, j - 2)
                            End If
                            If j = 4 Then
                                Cells(i + 1, j).Value = matriz(i, j - 2)
                            End If
                            If j = 5 Then
                                Cells(i + 1, j).Value = matriz(i, j + 1)
                            End If
                            If j = 6 Then
                                Cells(i + 1, j).Value = matriz(i, j + 1)
                            End If
                        Next j
                    Next i

End Sub
