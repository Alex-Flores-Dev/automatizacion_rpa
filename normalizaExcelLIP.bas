Set obj = createobject("Excel.Application")   'Creating an Excel Object
obj.visible=true                                    'Making an Excel Object visible
Set obj1 = obj.Workbooks.open("D:\AD\BUN\LIP\pruebasAlex.xlsx")    'Opening an Excel file
Set obj2=obj1.Worksheets("Completo Bs")    'Referring Sheet1 of excel file

Dim agencia , p , s , baja , glosa , semana , Original , columna , Final,i,j
Dim matriz() 
Dim vector() 

        varia = 8
        filas = 0
        columnas = 0
        
       ' Sheets("Completo Bs").Select
	   
	   obj1.Worksheets("Completo Bs").Select
       obj1.Worksheets("Completo Bs").Range("B8").Select
	
        obj2.Range("B8").Select
        Final = 0
        
        
            Do While Final <> "salir"
                    varia = varia + 1
                    filas = filas + 1
                If IsEmpty(obj2.Cells(varia, 2)) Then
                    alerta = alerta + 1
                Else
                    alerta = 0
                End If
                If alerta = 15 Then
                    Final = "salir"
                End If
            Loop
            columnas = 12
         ReDim matriz(filas + 8, columnas)
            
        For i = 8 To filas
            For j = 2 To columnas
                matriz(i, j) = obj2.Cells(i, j).Value
            Next 
        Next 


 obj1.Worksheets("Hoja1").Select
 obj1.Worksheets("Hoja1").Range("A2").Select
	
	
'obj2.Worksheets("Hoja1").Select
'Worksheets.Range("a2").Select

If matriz(10, 3) > 0 Then
    MsgBox matriz(10, 8)
    MsgBox columnas
End If
  
 For i = 8 To filas
    For j = 2 To columnas
        If matriz(i, 2) > 0 Then
                If IsNumeric(matriz(i, 2)) Then
                        obj2.ActiveCell.Offset(0, j - 2).Value = matriz(i, j)
                    Else
                        Exit For
                End If
            Else
                Exit For
        End If
    Next 
        If matriz(i, 2) > 0 Then
            If IsNumeric(matriz(i, 2)) Then
                obj2.ActiveCell.Offset(1, 0).Select
            End If
        End If
 Next 
 
 
 obj2.Range("C:C").EntireColumn.Delete
 obj2.Range("H:H").EntireColumn.Delete
 
  glosa = 3
  baja = 2
  
  
 For i = 8 To filas
    If matriz(i, 3) = "Glosa:" Then
        obj2.Cells(baja, 10).Value = matriz(i, 6)
            baja = baja + 1
    End If
    
    If matriz(i, 8) = "Datos Adicionales:" Then
        obj2.Cells(baja, 11).Value = matriz(i, 9)
    End If
 Next 

obj2.Cells(1, 1).Value = "Nro."
obj2.Cells(1, 2).Value = "Fecha"
obj2.Cells(1, 3).Value = "Hora"
obj2.Cells(1, 4).Value = "Tipo de Operacion"
obj2.Cells(1, 5).Value = "Participante Origen"
obj2.Cells(1, 6).Value = "Participante Destino"
obj2.Cells(1, 7).Value = "Retiro"
obj2.Cells(1, 8).Value = "Deposito"
obj2.Cells(1, 9).Value = "Cod"
obj2.Cells(1, 10).Value = "Glosa"
obj2.Cells(1, 11).Value = "Datos Adicionales"