Attribute VB_Name = "Modulo1"
Sub vouchers()
    
    'El codigo nombra hojas segun datos de la clumna A2 "Crono"
    'asigna ese dato al nombre de la hoja y a la casilla B17
    
    Dim i As Long, ultimaFila As Long
    Dim nombre As String
    Dim hojaCrono As Worksheet
    Dim hojaCustodia As Worksheet
    Dim nuevaHoja As Worksheet


    Set hojaCrono = ThisWorkbook.Worksheets("Crono")
    Set hojaCustodia = ThisWorkbook.Worksheets("Custodia")
    
        ' Verificar si hay hojas con números de carga y eliminarlas
    For i = Worksheets.Count To 1 Step -1
        carga = Worksheets(i).Name
        If IsNumeric(carga) Then
            Application.DisplayAlerts = False ' Desactivar alertas de confirmación
            Worksheets(carga).Delete
            Application.DisplayAlerts = True ' Activar alertas de confirmación
        End If
    Next i
    
    'Mostrar hoja "Custodia" si está oculta
    If hojaCustodia.Visible = xlSheetHidden Then
        hojaCustodia.Visible = xlSheetVisible
    End If
    
    'Obtener última fila de la columna A de la hoja "Crono"
    ultimaFila = hojaCrono.Cells(hojaCrono.Rows.Count, "A").End(xlUp).Row
    
    'Recorrer la columna A de la hoja "Crono" y crear una hoja para cada número de carga
    For i = 2 To ultimaFila
        nombre = hojaCrono.Cells(i, 1).Value
        
        'Verificar que el valor de la celda no está vacío
        If nombre <> "" Then
            'Crear una nueva hoja y renombrarla con el valor de la celda
            Set nuevaHoja = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
            nuevaHoja.Name = nombre
            
            'Copiar la información de la hoja "Custodia" a la nueva hoja y actualizar el valor de B17
            hojaCustodia.Range("B17").Value = nombre
            hojaCustodia.Cells.Copy nuevaHoja.Cells
            
            'Ocultar la hoja "Custodia" en lugar de eliminarla
            hojaCustodia.Visible = xlSheetHidden
        End If
    Next i
End Sub




