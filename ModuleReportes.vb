Imports System.Drawing.Printing
Imports Microsoft.Data.SqlClient

Module ModuloReportes

    ' Variables privadas para control de impresión PDF
    Private dgvAImprimir As DataGridView
    Private tituloReporteActual As String

    ''' <summary>
    ''' Muestra una vista previa rápida en un MessageBox con formato de ticket.
    ''' </summary>
    Public Sub ImprimirVistaPrevia(ByVal dgv As DataGridView, ByVal titulo As String)
        If dgv Is Nothing OrElse dgv.CurrentRow Is Nothing Then
            MsgBox("Seleccione un registro en la tabla para la vista previa.", MsgBoxStyle.Exclamation, "JosuOs")
            Return
        End If

        Dim fila = dgv.CurrentRow
        Dim r As String = ""
        Dim linea As String = "══════════════════════════════" & vbCrLf

        r &= "        AMERICAN COLLEGE         " & vbCrLf
        r &= "    SISTEMA ACADÉMICO JOSUOS     " & vbCrLf
        r &= linea
        r &= " FECHA: " & DateTime.Now.ToString("dd/MM/yyyy HH:mm") & vbCrLf
        r &= linea
        r &= "     " & titulo.ToUpper() & "       " & vbCrLf & linea

        ' Solo imprimimos las primeras 4 columnas relevantes para el ticket rápido
        For j As Integer = 0 To Math.Min(dgv.Columns.Count - 1, 4)
            Dim nombreCol As String = dgv.Columns(j).HeaderText.ToUpper()
            Dim valor As String = dgv.Rows(fila.Index).Cells(j).Value.ToString()

            ' Formateo de fecha si es necesario
            If nombreCol.Contains("FECHA") Then valor = CDate(valor).ToShortDateString()

            r &= " " & nombreCol.PadRight(10) & ": " & valor & vbCrLf
        Next

        r &= linea
        r &= "  Generado por Control Académico " & vbCrLf
        r &= "══════════════════════════════"

        MessageBox.Show(r, "Vista Previa de Ticket", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Genera un documento PDF profesional con diseño de tabla.
    ''' </summary>
    Public Sub ExportarAPDFGlobal(ByVal dgv As DataGridView, ByVal titulo As String)
        If dgv Is Nothing OrElse dgv.Rows.Count = 0 Then
            MsgBox("No hay registros activos para generar el PDF.", MsgBoxStyle.Exclamation)
            Return
        End If

        dgvAImprimir = dgv
        tituloReporteActual = titulo

        Dim pd As New PrintDocument()
        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF"

        ' --- CAMBIO A HORIZONTAL PARA MÁS ESPACIO ---
        pd.DefaultPageSettings.Landscape = True

        AddHandler pd.PrintPage, AddressOf ProcesoDibujoPDF

        Try
            pd.Print()
            MsgBox("PDF de '" & titulo & "' generado con éxito.", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Lógica de dibujo vectorial para el PDF (Evita que el texto se encime).
    ''' </summary>
    Private Sub ProcesoDibujoPDF(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        Dim fuenteTitulo As New Font("Arial", 18, FontStyle.Bold)
        Dim fuenteHeader As New Font("Arial", 8, FontStyle.Bold)
        Dim fuenteTexto As New Font("Arial", 8)

        ' Usamos el ancho real de la página horizontal
        Dim x As Integer = 50
        Dim y As Integer = 60
        Dim anchoTotal As Integer = e.PageBounds.Width - 100

        ' 1. ENCABEZADO
        e.Graphics.DrawString("AMERICAN COLLEGE - ACN LINUX", fuenteTitulo, Brushes.Firebrick, x, y)
        y += 35
        e.Graphics.DrawString("REPORTE OFICIAL: " & tituloReporteActual.ToUpper(), New Font("Arial", 10, FontStyle.Italic), Brushes.Black, x, y)
        y += 40

        ' 2. REPARTO DINÁMICO DE ESPACIO
        Dim numCols As Integer = dgvAImprimir.Columns.Count
        Dim anchoCol As Integer = anchoTotal / numCols

        ' 3. DIBUJAR ENCABEZADOS CON RECTÁNGULOS (Para que no choquen)
        e.Graphics.FillRectangle(Brushes.LightGray, x, y, anchoTotal, 25)
        For j As Integer = 0 To numCols - 1
            Dim rectHeader As New RectangleF(x + (j * anchoCol), y + 5, anchoCol - 5, 20)
            e.Graphics.DrawString(dgvAImprimir.Columns(j).HeaderText.ToUpper(), fuenteHeader, Brushes.Black, rectHeader)
        Next
        y += 30

        ' 4. DIBUJAR FILAS
        For i As Integer = 0 To dgvAImprimir.Rows.Count - 1
            If Not dgvAImprimir.Rows(i).IsNewRow Then
                For j As Integer = 0 To numCols - 1
                    If dgvAImprimir.Rows(i).Cells(j).Value IsNot Nothing Then
                        Dim texto As String = dgvAImprimir.Rows(i).Cells(j).Value.ToString()

                        ' Limpieza de fecha inteligente
                        If dgvAImprimir.Columns(j).HeaderText.ToUpper().Contains("FECHA") Then
                            Dim f As DateTime
                            If DateTime.TryParse(texto, f) Then texto = f.ToShortDateString()
                        End If

                        ' Dibujo con restricción de área (RectangleF)
                        Dim rectCelda As New RectangleF(x + (j * anchoCol), y, anchoCol - 5, 35) ' Más alto por si hay texto largo
                        e.Graphics.DrawString(texto, fuenteTexto, Brushes.Black, rectCelda)
                    End If
                Next
                e.Graphics.DrawLine(Pens.LightGray, x, y + 30, x + anchoTotal, y + 30)
                y += 35 ' Aumentamos el alto de la fila para que quepan nombres largos
            End If
            If y > e.MarginBounds.Bottom - 50 Then Exit For
        Next

        ' 5. PIE DE PÁGINA
        e.Graphics.DrawString("JosuOs Academia | " & DateTime.Now.ToString(), fuenteTexto, Brushes.Gray, x, e.MarginBounds.Bottom)
    End Sub

End Module
