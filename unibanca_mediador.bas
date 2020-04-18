Sub carga_unibanca()
Dim i, j, fn As Integer
    Dim cadena, ruta As String
    Dim arr
    Dim matriz() As Variant

' Para reconocer los path para unibanca y mediador
' la celda A1 siempre sera UNIBANCA
    ruta_unibanca = Worksheets("parametros").Cells(2, 1).Value

' la celda A2 siempre sera MEDIADOR
'    ruta_mediador= Worksheets("pametros").Cells(3, 1).Value

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
 '   ruta = "D:\AD\BUN\cuadre\procesado\2308\files\unibanca.txt"
' Cargar el archivo a la variable fn
    Open ruta_unibanca For Input As #fn

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
            Sheets("parametros").Delete
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

' la celda A2 siempre sera MEDIADOR
    ruta_mediador = Worksheets("parametros").Cells(3, 1).Value
' Cargar el archivo a la variable fn
    Open ruta_mediador For Input As #fn

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


