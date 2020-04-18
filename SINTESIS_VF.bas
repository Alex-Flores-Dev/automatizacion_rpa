Sub principal()
    Dim matriz(), matriz_final() As Variant
    Dim coef_llenar, servicio, arregla_filas, agencia, ciudad, departamento As Integer
    coef_llenar = 1
    servicio = 0
    arregla_filas = 0
    agencia = 1
    ciudad = 1
    departamento = 1
'eliminar alertas
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
'### DATA IMPORTANTE / DIMENSIONES DE LOS DOCUMENTOS
    yanbal = Array(56, 11, 15, 13, 12, 13)
    tigo_servicios_hogar = Array(42, 25, 15, 14, 13, 12)
    general = Array(41, 10, 15, 11, 13, 11)

' Integracion con las hojas
 ruta_principal = "C:\Users\Alex\Desktop\datos txt\NuevoSintesis\Sintesis Nuevo\"
    archivo = Dir(ruta_principal & "\*.txt")
    Do While Len(archivo) > 0
        ruta = ruta_principal & archivo

' para activar los libros virtuales
    Set HojaDestino = Workbooks(ThisWorkbook.Name).Worksheets("Hoja1")
    HojaDestino.Activate
    Worksheets.Add.Name = "Carga"
'Mandamos la ruta del archivo para analizar
    Call importar_datos(ruta, general)
' Analizamos el txt para saber que formato tendra
    evaluar = Range("A6").Value
    If InStr(1, evaluar, "YANBAL") Then
        formato = yanbal
    ElseIf InStr(1, evaluar, "TIGO SERVICIOS HOGAR") Then
        formato = tigo_servicios_hogar
    Else
        formato = general
    End If
    Sheets("Carga").Delete
    Worksheets.Add.Name = "Carga"
Call importar_datos(ruta, formato)

'Corregimos y definimos las dimensiones de la Matriz
 Rows(1).EntireRow.Delete
    Range("A1").Select
        limpiar = 1
        fila = 1
        Do While fila <= ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
                If InStr(1, Cells(fila, 1).Value, "------") > 0 Then
                        Rows(fila).EntireRow.Delete
                        fila = fila - 1
                    ElseIf InStr(1, Cells(fila, 1).Value, "     ") > 0 Then
                        Rows(fila).EntireRow.Delete
                        fila = fila - 1
                    ElseIf InStr(1, Cells(fila, 1).Value, "=======") > 0 Then
                        Rows(fila).EntireRow.Delete
                        fila = fila - 1
                    ElseIf InStr(1, Cells(fila, 1).Value, "*********") > 0 Then
                        Rows(fila).EntireRow.Delete
                        fila = fila - 1
                    ElseIf InStr(1, Cells(fila, 1).Value, "INTRA PLAT") > 0 Then
' Querido mantenedor de este código
' Una vez que hayas terminado de intentar "optimizar" esta rutina,
' y te hayas dado cuenta del terrible error que era hacerlo,
                        Rows(fila).EntireRow.Delete
                        Rows(fila).EntireRow.Delete
                        Rows(fila).EntireRow.Delete
                        Rows(fila).EntireRow.Delete
                        Rows(fila).EntireRow.Delete
                        fila = fila - 1
                    ElseIf InStr(1, Cells(fila, 1).Value, "@PB") > 0 Then
'SUPER MANUAL ESPERO REFACTORIZAR
                        Rows(fila).EntireRow.Delete
                        Rows(fila).EntireRow.Delete
                        Rows(fila).EntireRow.Delete
                        Rows(fila).EntireRow.Delete
                        Rows(fila).EntireRow.Delete
                        Rows(fila).EntireRow.Delete
                        fila = fila - 1
                    ElseIf Cells(fila, 1).Value = "SERVICIO" Then
                        Rows(fila).EntireRow.Delete
                        fila = fila - 1
                End If
                    fila = fila + 1
        Loop
        If Cells(fila - 1, 2).Value = "" Then
            Rows(fila - 1).EntireRow.Delete
        End If
        'SUPER MANUAL ESPERO REFACTORIZAR
        Columns("B").EntireColumn.Delete
        Columns("B").EntireColumn.Delete
        Columns("B").EntireColumn.Delete
        Columns("B").EntireColumn.Delete
' Dimensiones de la matriz
        filas = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
        columnas = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    ReDim matriz(filas, columnas)
' Cargamos datos a la matriz
        For i = 1 To filas
            For j = 1 To columnas
                matriz(i, j) = Cells(i, j).Value
            Next j
        Next i
' activamos hoja destino
    Sheets("Carga").Delete
    Set HojaDestino = Workbooks(ThisWorkbook.Name).Worksheets("Hoja1")
    HojaDestino.Activate
' Buscamos el DPTO



    compensa_filas = 0

'MAGICO NO TOCA POR NADA DEL MUNDO
        For i = 1 To filas
            j = 1
                If InStr(1, matriz(i, j), "TOTAL SERVICIO") > 0 Then
                    Do While coef_llenar <= servicio
                        Cells(coef_llenar, 4).Value = matriz(i, j)
                        coef_llenar = coef_llenar + 1
                    Loop
                        compensa_filas = compensa_filas + 1

                ElseIf InStr(1, matriz(i, j), "TOTAL AGENCIA") > 0 Then
                    Do While agencia <= servicio
                        Cells(agencia, 3).Value = matriz(i, j)
                        agencia = agencia + 1
                    Loop
                        compensa_filas = compensa_filas + 1

                ElseIf InStr(1, matriz(i, j), "TOTAL CIUDAD") > 0 Then
                    Do While ciudad <= servicio
                        Cells(ciudad, 2).Value = matriz(i, j)
                        ciudad = ciudad + 1
                    Loop
                        compensa_filas = compensa_filas + 1
                ElseIf InStr(1, matriz(i, j), "TOTAL DPTO") > 0 Then
                    Do While departamento <= servicio
                        Cells(departamento, 1).Value = matriz(i, j)
                        departamento = departamento + 1
                    Loop
                        compensa_filas = compensa_filas + 1
                Else
                    Cells(i - compensa_filas + arregla_filas, 5).Value = matriz(i, j)
                    Cells(i - compensa_filas + arregla_filas, 6).Value = matriz(i, j + 1)
                    Cells(i - compensa_filas + arregla_filas, 7).Value = matriz(i, j + 2)
                    servicio = servicio + 1

                End If
        Next i
        'ELIMINAMOS LA FILA DE TOTALES
        Rows(servicio).EntireRow.Delete
        'COMPENSA LOS CAMPOS ELIMINADOS
        servicio = servicio - 1
        'PARA QUE LA MATRIZ NO SE COMPLIQUE
        arregla_filas = servicio
        ' CAMBIA AL SIGUIENTE ARCHIVO
        archivo = Dir()
    Loop
    ' ULTIMA LIMPIEZA DE CAMPOS VACIOS SOLO BUSCAMOS CAMPOS VACIOS Y ELIMINAMOS
        last_clean = 1
            Do While last_clean <= ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
                If IsEmpty(Cells(last_clean, 6)) Then
                    Rows(last_clean).EntireRow.Delete
                    last_clean = last_clean - 1
                End If
                last_clean = last_clean + 1
            Loop
'LE DAMOS FORMATO GENERAL PORQ ALGUNAS CELDAS ESTAN CON OTROS FORMATOS
        Range("F:F").Select
        Selection.NumberFormat = "General"
        Range("G:G").Select
        Selection.NumberFormat = "General"
        Range("a1").Select
End Sub

'Puede que pienses que sabes qué hace el siguiente código
'Pero no lo sabes. Confía en mi.

Sub importar_datos(ruta, general)
    tipo = "TEXT"
    ruta_completa = tipo & ";" & ruta
' Esta parte se utilizo para jalar el txt y darle forma porq el archivo no tiene separador ";"
    With ActiveSheet.QueryTables.Add(Connection:= _
        ruta_completa _
        , Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlMSDOS
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1)
        .TextFileFixedColumnWidths = general
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery = False
    End With
'    ActiveWindow.SmallScroll Down = -18
End Sub





