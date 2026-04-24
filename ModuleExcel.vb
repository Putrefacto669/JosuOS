Imports Microsoft.Office.Interop

Module ModuloExcel
    Public Sub ExportarDataGridViewAExcel(ByVal dgv As DataGridView, ByVal nombreHoja As String)
        ' Validamos si hay algo que exportar
        If dgv.Rows.Count = 0 Then
            MsgBox("No hay datos en la tabla para exportar a Excel.", MsgBoxStyle.Exclamation)
            Return
        End If

        Try
            Dim excelApp As New Excel.Application
            Dim libro As Excel.Workbook = excelApp.Workbooks.Add()
            Dim hoja As Excel.Worksheet = libro.Sheets(1)
            hoja.Name = nombreHoja

            ' 1. Exportar los Encabezados (Nombres de las columnas)
            For i As Integer = 1 To dgv.Columns.Count
                hoja.Cells(1, i) = dgv.Columns(i - 1).HeaderText
                hoja.Cells(1, i).Font.Bold = True ' Poner en negrita el título
            Next

            ' 2. Exportar las Filas y Celdas
            For i As Integer = 0 To dgv.Rows.Count - 1
                For j As Integer = 0 To dgv.Columns.Count - 1
                    ' Validamos que la celda no sea nula para evitar errores
                    If dgv.Rows(i).Cells(j).Value IsNot Nothing Then
                        hoja.Cells(i + 2, j + 1) = dgv.Rows(i).Cells(j).Value.ToString()
                    End If
                Next
            Next

            ' 3. Ajustar automáticamente el ancho de las columnas
            hoja.Columns.AutoFit()

            ' 4. Hacer visible el Excel al terminar
            excelApp.Visible = True
            MsgBox("¡Datos exportados a Excel con éxito!", MsgBoxStyle.Information, "JosuOs - Éxito")

        Catch ex As Exception
            MsgBox("Error al exportar: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
End Module
