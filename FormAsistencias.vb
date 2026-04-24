Imports Microsoft.Data.SqlClient
Imports System.Drawing.Printing
Public Class FormAsistencias
    Private Sub FormAsistencias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 1. Configurar estados
        cbEstado.Items.Clear()
        cbEstado.Items.AddRange(New String() {"Presente", "Falta", "Justificado"})
        cbEstado.SelectedIndex = 0

        ' 2. Cargar datos
        CargarMateriasAsignadas()
        ListarAsistencias()
        ' En el FormAsistencias_Load agrega:
        cmbCuatrimestre.Items.Clear()
        cmbCuatrimestre.Items.AddRange(New String() {"Todos", "I Cuatrimestre", "II Cuatrimestre", "III Cuatrimestre"})
        cmbCuatrimestre.SelectedIndex = 0

    End Sub
    Private Sub CargarMateriasAsignadas()
        Try
            Conectar()
            ' Usamos GROUP BY para colapsar los duplicados en una sola fila
            ' Traemos el MIN(IdMateria) para tener un ID de referencia único
            Dim sql As String = "SELECT MIN(M.IdMateria) as IdMateria, " &
                               "(M.NombreMateria + ' - Prof. ' + D.NombreFull) as Info " &
                               "FROM Materias M " &
                               "INNER JOIN Docentes D ON M.IdDocente = D.IdDocente " &
                               "GROUP BY M.NombreMateria, D.NombreFull"

            Dim da As New SqlDataAdapter(sql, cnn)
            Dim dt As New DataTable
            da.Fill(dt)

            ' Limpiamos antes de asignar
            cbMaterias.DataSource = Nothing
            cbMaterias.Items.Clear()

            cbMaterias.DataSource = dt
            cbMaterias.DisplayMember = "Info"
            cbMaterias.ValueMember = "IdMateria"

            ' Configuramos el buscador tipo Google
            cbMaterias.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            cbMaterias.AutoCompleteSource = AutoCompleteSource.ListItems

        Catch ex As Exception
            MsgBox("Error al limpiar duplicados: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub ListarAsistencias()
        Try
            Conectar()
            Dim sql As String = "SELECT A.IdAsistencia, M.NombreMateria, D.NombreFull as Docente, A.Fecha, A.Estado " &
                               "FROM Asistencias A " &
                               "INNER JOIN Materias M ON A.IdMateria = M.IdMateria " &
                               "INNER JOIN Docentes D ON M.IdDocente = D.IdDocente " &
                               "ORDER BY A.Fecha DESC"
            Dim da As New SqlDataAdapter(sql, cnn)
            Dim dt As New DataTable
            da.Fill(dt)
            dgvAsistencias.DataSource = dt
        Catch ex As Exception
            MsgBox("Error al listar: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Try
            If cbMaterias.SelectedValue Is Nothing Then Return

            Conectar()
            Dim sql As String = "INSERT INTO Asistencias (IdMateria, Fecha, Estado) VALUES (@idM, @fec, @est)"
            Dim cmd As New SqlCommand(sql, cnn)
            cmd.Parameters.AddWithValue("@idM", cbMaterias.SelectedValue)
            cmd.Parameters.AddWithValue("@fec", dtpFecha.Value.Date)
            cmd.Parameters.AddWithValue("@est", cbEstado.Text)

            cmd.ExecuteNonQuery()
            MsgBox("Asistencia registrada correctamente en JosuOs", MsgBoxStyle.Information)
            ListarAsistencias()
        Catch ex As Exception
            MsgBox("Error al guardar asistencia: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If cbMaterias.SelectedValue Is Nothing Then Return

            Conectar()

            ' 1. Consultamos el Total de Horas (Créditos)
            Dim sqlHoras As String = "SELECT SUM(M.Creditos) FROM Asistencias A " &
                                    "INNER JOIN Materias M ON A.IdMateria = M.IdMateria " &
                                    "WHERE A.IdMateria = @idM AND A.Estado = 'Presente'"

            Dim cmdH As New SqlCommand(sqlHoras, cnn)
            cmdH.Parameters.AddWithValue("@idM", cbMaterias.SelectedValue)
            Dim totalHoras = cmdH.ExecuteScalar()

            ' 2. Consultamos Datos del Docente para la "Colilla" (Nivel y Facultad)
            ' Ajustamos el SQL para traer la info del docente dueño de esa materia
            Dim sqlDoc As String = "SELECT D.NivelAcademico, D.Facultad FROM Docentes D " &
                                  "INNER JOIN Materias M ON D.IdDocente = M.IdDocente " &
                                  "WHERE M.IdMateria = @idM"

            Dim cmdD As New SqlCommand(sqlDoc, cnn)
            cmdD.Parameters.AddWithValue("@idM", cbMaterias.SelectedValue)
            Dim reader = cmdD.ExecuteReader()

            Dim nivel As String = "No especificado"
            Dim facultad As String = "Ingeniería" ' Valor por defecto

            If reader.Read() Then
                nivel = reader("NivelAcademico").ToString()
                facultad = reader("Facultad").ToString()
            End If
            reader.Close()

            ' 3. Verificamos y Mostramos el Reporte Estilo American College
            If IsDBNull(totalHoras) Or totalHoras Is Nothing Then
                MessageBox.Show("No se registran horas impartidas para esta materia.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                Dim reporte As String = "          AMERICAN COLLEGE" & vbCrLf &
                                        "   REPORTE DE CUMPLIMIENTO DOCENTE" & vbCrLf &
                                        "------------------------------------------" & vbCrLf &
                                        "Nivel: " & nivel & vbCrLf &
                                        "Facultad: " & facultad & vbCrLf &
                                        "Materia: " & cbMaterias.Text & vbCrLf &
                                        "------------------------------------------" & vbCrLf &
                                        "TOTAL HORAS IMPARTIDAS: " & totalHoras.ToString() & " hrs" & vbCrLf &
                                        "------------------------------------------" & vbCrLf &
                                        "Estado: Verificado por Control Académico"

                MessageBox.Show(reporte, "JosuOs - ACN LINUX", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MsgBox("Error en el reporte: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        ' Lógica similar a los otros formularios para borrar una asistencia mal marcada
        If dgvAsistencias.CurrentRow Is Nothing Then Return

        Dim id As Integer = dgvAsistencias.CurrentRow.Cells(0).Value
        If MsgBox("¿Desea eliminar este registro de asistencia?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Try
                Conectar()
                Dim cmd As New SqlCommand("DELETE FROM Asistencias WHERE IdAsistencia = @id", cnn)
                cmd.Parameters.AddWithValue("@id", id)
                cmd.ExecuteNonQuery()
                ListarAsistencias()
            Catch ex As Exception
                MsgBox(ex.Message)
            Finally
                Desconectar()
            End Try
        End If
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        ExportarDataGridViewAExcel(dgvAsistencias, "Reporte de Asistencias")
    End Sub

    Private Sub btnTodos_Click(sender As Object, e As EventArgs) Handles btnTodos.Click
        Try
            Conectar()
            Dim sql As String = "SELECT D.NombreFull AS Docente, M.Facultad, M.Turno, M.Cuatrimestre, " &
                           "SUM(M.Creditos) AS [Total Horas], " &
                           "(SELECT COUNT(*) FROM Examenes E WHERE E.NombreDocente = D.NombreFull) AS [Examenes] " &
                           "FROM Docentes D " &
                           "INNER JOIN Materias M ON D.IdDocente = M.IdDocente " &
                           "INNER JOIN Asistencias A ON M.IdMateria = A.IdMateria " &
                           "WHERE A.Estado = 'Presente' AND A.Fecha BETWEEN @inicio AND @fin "

            If cmbCuatrimestre.Text <> "Todos" Then
                sql &= "AND M.Cuatrimestre = @cuat "
            End If

            sql &= "GROUP BY D.NombreFull, M.Facultad, M.Turno, M.Cuatrimestre"

            Dim cmd As New SqlCommand(sql, cnn)
            cmd.Parameters.AddWithValue("@inicio", DtpInicio.Value.Date)
            cmd.Parameters.AddWithValue("@fin", DtpFin.Value.Date)
            If cmbCuatrimestre.Text <> "Todos" Then cmd.Parameters.AddWithValue("@cuat", cmbCuatrimestre.Text)

            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            DgvTotales.DataSource = dt ' Aquest és el grid de baix de la teva imatge
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub btnCalcularColillas_Click(sender As Object, e As EventArgs) Handles btnCalcularColillas.Click
        Try
            If cbMaterias.Text = "" Then Return
            Conectar()

            ' 1. Limpieza de datos: Extraemos el nombre del docente quitando el "Prof."
            Dim partes = cbMaterias.Text.Split("-")
            Dim nombreDocente As String = ""
            If partes.Length > 1 Then
                nombreDocente = partes(1).Replace("Prof.", "").Trim()
            End If

            ' 2. Consulta SQL: Agrupamos por Materia para el desglose detallado
            Dim sql As String = "SELECT M.NombreMateria AS Materia, M.Facultad, M.Turno, " &
                   "SUM(M.Creditos) AS [Horas], " &
                   "(SELECT COUNT(*) FROM Examenes E " &
                   " WHERE E.NombreDocente LIKE @doc " & ' Filtramos por docente
                   " AND E.NombreClase LIKE '%' + M.NombreMateria + '%') AS [Examenes] " & ' Filtro flexible por materia
                   "FROM Asistencias A " &
                   "INNER JOIN Materias M ON A.IdMateria = M.IdMateria " &
                   "INNER JOIN Docentes D ON M.IdDocente = D.IdDocente " &
                   "WHERE D.NombreFull LIKE @doc " &
                   "AND A.Estado LIKE '%Presente%' " &
                   "AND CAST(A.Fecha AS DATE) BETWEEN @ini AND @fin " &
                   "GROUP BY M.NombreMateria, M.Facultad, M.Turno"

            Dim cmd As New SqlCommand(sql, cnn)
            cmd.Parameters.AddWithValue("@doc", "%" & nombreDocente & "%")
            cmd.Parameters.Add("@ini", SqlDbType.Date).Value = DtpInicio.Value.Date
            cmd.Parameters.Add("@fin", SqlDbType.Date).Value = DtpFin.Value.Date

            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            ' Mostramos los datos en el DataGridView
            DgvTotales.DataSource = dt

            If dt.Rows.Count > 0 Then
                ' 3. Construcción del Mensaje de la Colilla
                Dim totalHoras As Integer = 0
                Dim totalExamenes As Integer = 0
                Dim detalleMaterias As String = ""

                For Each fila As DataRow In dt.Rows
                    Dim materia = fila("Materia").ToString()
                    Dim horas = CInt(fila("Horas"))
                    Dim examenes = CInt(fila("Examenes"))

                    detalleMaterias &= $"• {materia}: {horas} hrs | {examenes} exam." & vbCrLf
                    totalHoras += horas
                    totalExamenes += examenes
                Next

                ' El "Ticket" o Colilla final
                Dim colilla As String =
                    "==========================================" & vbCrLf &
                    "       AMERICAN COLLEGE - JOSUOS        " & vbCrLf &
                    "         REPORTE DE ASISTENCIA          " & vbCrLf &
                    "==========================================" & vbCrLf &
                    "FECHA: " & DateTime.Now.ToShortDateString() & vbCrLf &
                    "DOCENTE: " & nombreDocente.ToUpper() & vbCrLf &
                    "PERIODO: " & DtpInicio.Value.ToShortDateString() & " al " & DtpFin.Value.ToShortDateString() & vbCrLf &
                    "------------------------------------------" & vbCrLf &
                    "DETALLE POR MATERIA:" & vbCrLf &
                    detalleMaterias &
                    "------------------------------------------" & vbCrLf &
                    "TOTAL HORAS IMPARTIDAS: " & totalHoras & " hrs" & vbCrLf &
                    "TOTAL EXÁMENES REALIZADOS: " & totalExamenes & vbCrLf &
                    "==========================================" & vbCrLf &
                    "ESTADO: Válido para trámite de colilla"

                MsgBox(colilla, MsgBoxStyle.Information, "Colilla de Pago Generada")
            Else
                MsgBox("No se encontraron registros de asistencia 'Presente' para " & nombreDocente & " en las fechas seleccionadas.", MsgBoxStyle.Exclamation)
            End If

        Catch ex As Exception
            MsgBox("Error al generar colilla: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub btnExportar_Click(sender As Object, e As EventArgs) Handles btnExportar.Click
        ExportarDataGridViewAExcel(DgvTotales, "Reporte Detallado de Asistencias")
    End Sub

    Private Sub btnEliminar2_Click(sender As Object, e As EventArgs) Handles btnEliminar2.Click
        If DgvTotales.SelectedRows.Count = 0 Then
            MsgBox("Selecciona una fila en la tabla de abajo para eliminar.", MsgBoxStyle.Exclamation)
            Return
        End If

        Dim respuesta As DialogResult = MessageBox.Show("¿Seguro que deseas eliminar los registros de esta materia?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If respuesta = DialogResult.Yes Then
            Try
                Conectar()
                Dim materiaAEliminar As String = DgvTotales.CurrentRow.Cells("Materia").Value.ToString()

                ' CAMBIO CLAVE: Usamos IN en lugar de = para evitar el error de múltiples valores
                Dim sql As String = "DELETE FROM Asistencias WHERE IdMateria IN " &
                               "(SELECT IdMateria FROM Materias WHERE NombreMateria = @nom) " &
                               "AND Estado = 'Presente' AND CAST(Fecha AS DATE) BETWEEN @ini AND @fin"

                Dim cmd As New SqlCommand(sql, cnn)
                cmd.Parameters.AddWithValue("@nom", materiaAEliminar)
                cmd.Parameters.Add("@ini", SqlDbType.Date).Value = DtpInicio.Value.Date
                cmd.Parameters.Add("@fin", SqlDbType.Date).Value = DtpFin.Value.Date

                Dim filasAfetadas As Integer = cmd.ExecuteNonQuery()

                If filasAfetadas > 0 Then
                    MsgBox("Registros eliminados correctamente. Por favor, vuelve a dar clic en 'Calcular Colilla' para actualizar.", MsgBoxStyle.Information)
                    DgvTotales.DataSource = Nothing
                Else
                    MsgBox("No se encontraron registros para eliminar en esas fechas.", MsgBoxStyle.Information)
                End If

            Catch ex As Exception
                MsgBox("Error al eliminar: " & ex.Message)
            Finally
                Desconectar()
            End Try
        End If
    End Sub
    Private WithEvents PD As New PrintDocument
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Configurar el diálogo de impresión para que el usuario elija "Microsoft Print to PDF"
        Dim printDlg As New PrintDialog()
        printDlg.Document = PD

        If printDlg.ShowDialog() = DialogResult.OK Then
            PD.Print() ' Esto dispara el evento PrintPage de abajo
        End If
    End Sub

    ' Este es el "corazón" del PDF, aquí dibujamos el diseño manualmente
    Private Sub PD_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PD.PrintPage
        Dim g As Graphics = e.Graphics
        Dim fuenteTitulo As New Font("Arial", 16, FontStyle.Bold)
        Dim fuenteSub As New Font("Arial", 12, FontStyle.Bold)
        Dim fuenteTexto As New Font("Arial", 10)
        Dim fuenteBold As New Font("Arial", 10, FontStyle.Bold)

        Dim y As Integer = 50 ' Posición vertical inicial
        Dim x As Integer = 50 ' Margen izquierdo

        ' 1. Encabezado
        g.DrawString("AMERICAN COLLEGE", fuenteTitulo, Brushes.Black, x + 200, y)
        y += 30
        g.DrawString("SISTEMA DE CONTROL ACADÉMICO - JOSUOS", fuenteSub, Brushes.Black, x + 120, y)
        y += 40
        g.DrawLine(Pens.Black, x, y, 750, y) ' Línea divisoria
        y += 10

        ' 2. Datos del Reporte
        g.DrawString("REPORTE CONSOLIDADO DE COLILLAS", fuenteBold, Brushes.Black, x, y)
        y += 20
        g.DrawString("Rango: " & DtpInicio.Value.ToShortDateString & " al " & DtpFin.Value.ToShortDateString, fuenteTexto, Brushes.Black, x, y)
        y += 20
        g.DrawString("Cuatrimestre: " & cmbCuatrimestre.Text, fuenteTexto, Brushes.Black, x, y)
        y += 30

        ' 3. Encabezados de la Tabla (Dibujamos las celdas)
        Dim rectHeader As New Rectangle(x, y, 700, 25)
        g.FillRectangle(Brushes.LightGray, rectHeader)
        g.DrawRectangle(Pens.Black, rectHeader)

        ' Dibujamos los nombres de las columnas
        g.DrawString("Docente", fuenteBold, Brushes.Black, x + 5, y + 5)
        g.DrawString("Facultad", fuenteBold, Brushes.Black, x + 200, y + 5)
        g.DrawString("Turno", fuenteBold, Brushes.Black, x + 400, y + 5)
        g.DrawString("Horas", fuenteBold, Brushes.Black, x + 500, y + 5)
        g.DrawString("Exámenes", fuenteBold, Brushes.Black, x + 600, y + 5)

        y += 25

        ' 4. Filas de Datos (recorremos el DataGridView dgvTotales)
        For Each fila As DataGridViewRow In DgvTotales.Rows
            If Not fila.IsNewRow Then
                ' Dibujamos cada celda
                g.DrawString(fila.Cells("Docente").Value.ToString(), fuenteTexto, Brushes.Black, x + 5, y + 5)
                g.DrawString(fila.Cells("Facultad").Value.ToString(), fuenteTexto, Brushes.Black, x + 200, y + 5)
                g.DrawString(fila.Cells("Turno").Value.ToString(), fuenteTexto, Brushes.Black, x + 400, y + 5)
                g.DrawString(fila.Cells("Total Horas").Value.ToString(), fuenteTexto, Brushes.Black, x + 500, y + 5)
                g.DrawString(fila.Cells("Examenes").Value.ToString(), fuenteTexto, Brushes.Black, x + 600, y + 5)

                g.DrawLine(Pens.LightGray, x, y + 25, 750, y + 25) ' Línea tenue entre filas
                y += 25

                ' Control de salto de página (opcional)
                If y > 1000 Then
                    e.HasMorePages = True
                    Return
                End If
            End If
        Next

        ' 5. Pie de página y firmas
        y += 50
        g.DrawLine(Pens.Black, x + 50, y, x + 250, y)
        g.DrawLine(Pens.Black, x + 450, y, x + 650, y)
        y += 5
        g.DrawString("Firma Control Académico", fuenteTexto, Brushes.Black, x + 70, y)
        g.DrawString("Sello de Facultad", fuenteTexto, Brushes.Black, x + 500, y)

        e.HasMorePages = False ' No hay más páginas
    End Sub
End Class
