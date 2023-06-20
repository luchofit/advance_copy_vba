Sub CopiarValores()
    Dim Original As Workbook
    Dim ExternoCopy As Workbook
    Dim wsOriginal As Worksheet
    Dim destCell As Range
    Dim destColumn As String
    Dim columnas() As Variant
    Dim wsExternoCopy As Worksheet
    Dim InitialSheet As Worksheet
    Dim isFirstSheet As Boolean
    Dim previousSheet As Worksheet
    Dim columna As Variant
    Dim lastRow As Long
    Dim lastRow_old As Long
    Dim foundColumn As Range
    Dim numSelectedCells As Long
    Dim selectedRange As Variant
    
    ' Abre los archivos
    Set Original = ThisWorkbook
    Set ExternoCopy = Workbooks.Open("C:\Users\castrol2\Downloads\Archivo extraccion.xlsx") ' Archivo Copy
    
        ' Comprueba si la hoja de cálculo existe
    On Error Resume Next
    Set wsOriginal = Original.Worksheets("Referencias")
    On Error GoTo 0
    
        If wsOriginal Is Nothing Then
        MsgBox "No se encontró la hoja de cálculo 'Reference' en el archivo Hot Spot."
        wbCopy.Close SaveChanges:=False
        Exit Sub
    End If
    
    'Punto de partida de la copiaxion
    Set destCell = wsOriginal.Range("A2")
    destColumn = "A" ' Columna inicial
    
        ' Define los nombres de las columnas a copiar
    columnas = Array("ID", "PRODUCTO", "CODIGO")
    
    
        ' Itera sobre las pestañas del archivo Copy
    For i = 1 To ExternoCopy.Worksheets.Count
        Set wsExternoCopy = ExternoCopy.Worksheets(i)
        Set InitialSheet = ExternoCopy.Worksheets("Sheet1")
        isFirstSheet = (i = 1)
        If i = 1 And InitialSheet.Name = "Sheet1" Then
        Set previousSheet = ExternoCopy.Worksheets(i)
        Else
        Set previousSheet = ExternoCopy.Worksheets(i - 1)
        End If
        
                ' Copia los valores de cada columna en el archivo original
        For Each columna In columnas
            ' Busca la columna correspondiente en Copy
            Set foundColumn = wsExternoCopy.Range("A1:E1").Find(columna, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
            
            If Not foundColumn Is Nothing Then
                ' Verifica si hay valores en la columna encontrada
                If Application.WorksheetFunction.CountA(wsExternoCopy.Columns(foundColumn.Column)) > 1 Then
                ' Encuentra la última fila en el archivo original
                If Not isFirstSheet And wsExternoCopy.Name <> previousSheet.Name Then
                    If columna <> "PRODUCTO" And columna <> "CODIGO" Then
                        lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row
                    Else
                    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row
                    lastRow = lastRow - numSelectedCells
                    End If
            Else
                lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row
                lastRow = lastRow - numSelectedCells
            End If
            
                    ' Copia los valores desde Copy a Hot Spot
                    
                    Set selectedRange = wsExternoCopy.Range(foundColumn, wsExternoCopy.Cells(wsExternoCopy.Rows.Count, foundColumn.Column).End(xlUp)).Offset(1)
                    ' Obtener el número de celdas seleccionadas
                    numSelectedCells = selectedRange.Count
                    selectedRange.Copy
                    wsOriginal.Cells(lastRow + 1, destCell.Column).PasteSpecial Paste:=xlPasteValues
                    
                    ' Ajusta la celda de destino para la próxima iteración
                    If destCell.Column < wsOriginal.Cells(1, "C").Column Then
                        Set destCell = destCell.Offset(, 1)
                    Else
                        Set destCell = wsOriginal.Range("A2")
                    End If
                    
                Else
                    MsgBox "La columna '" & columna & "' en la hoja de cálculo " & wsCopy.Name & " no contiene valores."
                End If
            Else
                MsgBox "No se encontró la columna '" & columna & "' en la hoja de cálculo " & wsCopy.Name & "."
            End If
        
        Next columna
    Next i
  
    ' Cierra el archivo Copy sin guardar cambios
    wbCopy.Close SaveChanges:=False
    
    ' Liberar memoria
    Set wsOriginal = Nothing
    Set wsExternoCopy = Nothing
    Set Original = Nothing
    Set ExternoCopy = Nothing
    
    MsgBox "¡Valores copiados correctamente!"

End Sub

