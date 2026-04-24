Imports Microsoft.Data.SqlClient
Imports Windows.Win32
Imports System.IO
Imports System.Diagnostics
Public Class FormMaterias
    ' Variable para evitar que el filtro se dispare mientras se carga el combo inicial
    Private cargandoInicial As Boolean = True
    Private Sub FormMaterias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cargandoInicial = True

        ' 1. Configurar ComboBox de Horarios
        cbHorario.Items.Clear()
        cbHorario.Items.AddRange(New String() {"8:00 – 9:20", "9:30 – 10:50", "11:10 – 12:20", "1:00 – 2:20"})
        cbHorario.SelectedIndex = 0

        ' 2. ACTIVAR BÚSQUEDA FLUIDA 
        cbDocente.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        cbDocente.AutoCompleteSource = AutoCompleteSource.ListItems

        ' 3. Cargar datos base
        CargarDocentesCombo()
        ListarAsignaciones()

        cargandoInicial = False
        ' Disparamos el primer filtro manualmente
        If cbDocente.Items.Count > 0 Then cbDocente_SelectedIndexChanged(Nothing, Nothing)

        cbTurno.Items.Clear()
        cbTurno.Items.AddRange(New String() {"Matutino", "Vespertino", "Sabatino", "Dominical"})
        cbTurno.SelectedIndex = 0

        ' 4. Configurar ComboBox de Días
        cbDia.Items.Clear()
        cbDia.Items.AddRange(New String() {"Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"})
        cbDia.SelectedIndex = 0

        cmbCuatrimestre.Items.Clear()
        cmbCuatrimestre.Items.AddRange(New String() {"I Cuatrimestre", "II Cuatrimestre", "III Cuatrimestre"})
        cmbCuatrimestre.SelectedIndex = 0

        txtFacultad.ReadOnly = True ' La facultat ve del docent
        Try
            Conectar()
            Dim sql As String = "SELECT NombreFull FROM Docentes ORDER BY NombreFull ASC"
            Dim da As New SqlDataAdapter(sql, cnn)
            Dim dt As New DataTable
            da.Fill(dt)

            ' --- AQUÍ ESTÁ EL TRUCO PARA EVITAR EL ERROR ---
            ' Creamos una fila nueva manualmente en la tabla de datos
            Dim filaTodos As DataRow = dt.NewRow()
            filaTodos("NombreFull") = "--- TODOS LOS DOCENTES ---"
            ' La insertamos en la posición 0 (al principio de la lista)
            dt.Rows.InsertAt(filaTodos, 0)

            ' Ahora sí, vinculamos la tabla completa al ComboBox
            cbDocente.DataSource = dt
            cbDocente.DisplayMember = "NombreFull"

            ' Para que empiece vacío y no seleccione a nadie por defecto
            cbDocente.SelectedIndex = -1

        Catch ex As Exception
            MsgBox("Error al cargar docentes: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub
    Private Sub CargarDocentesCombo()
        Try
            Conectar()
            Dim dt As New DataTable
            Dim da As New SqlDataAdapter("SELECT IdDocente, NombreFull FROM Docentes", cnn)
            da.Fill(dt)

            cbDocente.DataSource = dt
            cbDocente.DisplayMember = "NombreFull" ' Lo que se muestra al usuario
            cbDocente.ValueMember = "IdDocente"    ' El ID real
        Catch ex As Exception
            MsgBox("Error al cargar docentes: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub
    Private Sub ListarAsignaciones()
        Try
            Conectar()
            Dim sql As String = "SELECT M.IdMateria, M.NombreMateria, M.DiaSemana, M.Horario, M.Turno, " &
                           "D.NombreFull as Docente, M.Creditos " &
                           "FROM Materias M INNER JOIN Docentes D ON M.IdDocente = D.IdDocente"
            Dim da As New SqlDataAdapter(sql, cnn)
            Dim dt As New DataTable
            da.Fill(dt)
            DgvMaterias.DataSource = dt
        Catch ex As Exception
            MsgBox("Error al listar: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Try
            Conectar()
            Dim sql As String = "INSERT INTO Materias (NombreMateria, Creditos, IdDocente, Horario, DiaSemana, Turno, Cuatrimestre, Facultad) " &
                           "VALUES (@nom, @cre, @idD, @hor, @dia, @turno, @cuat, @fac)"
            Dim cmd As New SqlCommand(sql, cnn)
            cmd.Parameters.AddWithValue("@nom", cbMateria.Text)
            cmd.Parameters.AddWithValue("@cre", NumericUpDown1.Value)
            cmd.Parameters.AddWithValue("@idD", cbDocente.SelectedValue)
            cmd.Parameters.AddWithValue("@hor", cbHorario.Text)
            cmd.Parameters.AddWithValue("@dia", cbDia.Text)
            cmd.Parameters.AddWithValue("@turno", cbTurno.Text)
            cmd.Parameters.AddWithValue("@cuat", cmbCuatrimestre.Text)
            cmd.Parameters.AddWithValue("@fac", txtFacultad.Text)

            cmd.ExecuteNonQuery()
            MsgBox("Materia registrada en JosuOs con éxito")
            ListarAsignaciones()
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If DgvMaterias.SelectedRows.Count = 0 Then
            MsgBox("Seleccione una fila para eliminar")
            Return
        End If

        Dim id As Integer = DgvMaterias.CurrentRow.Cells(0).Value
        If MsgBox("¿Eliminar esta asignación?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.Yes Then
            Try
                Conectar()
                Dim cmd As New SqlCommand("DELETE FROM Materias WHERE IdMateria = @id", cnn)
                cmd.Parameters.AddWithValue("@id", id)
                cmd.ExecuteNonQuery()
                ListarAsignaciones()
            Catch ex As Exception
                MsgBox("Error: " & ex.Message)
            Finally
                Desconectar()
            End Try
        End If
    End Sub

    Private Sub cbDocente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbDocente.SelectedIndexChanged
        If cbDocente.SelectedValue IsNot Nothing AndAlso IsNumeric(cbDocente.SelectedValue) Then
            ActualizarDatosDocente(CInt(cbDocente.SelectedValue))
        End If
    End Sub
    Private Sub ActualizarDatosDocente(id As Integer)
        Try
            Conectar()
            Dim cmd As New SqlCommand("SELECT Especialidad, Facultad FROM Docentes WHERE IdDocente = @id", cnn)
            cmd.Parameters.AddWithValue("@id", id)
            Dim reader = cmd.ExecuteReader()
            If reader.Read() Then
                txtEspecialidad.Text = reader("Especialidad").ToString()
                txtFacultad.Text = reader("Facultad").ToString()
                FiltrarMateriasPorEspecialidad(reader("Especialidad").ToString())
            End If
            reader.Close()
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub ActualizarInterfazPorDocente(ByVal idDoc As Integer)
        Try
            Conectar()
            ' Buscamos la especialidad en la tabla Docentes
            Dim cmd As New SqlCommand("SELECT Especialidad FROM Docentes WHERE IdDocente = @id", cnn)
            cmd.Parameters.AddWithValue("@id", idDoc)

            Dim especialidad As String = cmd.ExecuteScalar()?.ToString()

            If Not String.IsNullOrEmpty(especialidad) Then
                ' 1. El campo especialidad se llena solo (Automatizado)
                txtEspecialidad.Text = especialidad

                ' 2. Filtramos las materias que ese docente puede dar (Filtro inteligente)
                FiltrarMateriasPorEspecialidad(especialidad)
            End If
        Catch ex As Exception
            ' Error silencioso
        Finally
            Desconectar()
        End Try
    End Sub
    Private Sub FiltrarMateriasPorEspecialidad(ByVal esp As String)
        cbMateria.Items.Clear()

        ' Relacionamos especialidad con asignaturas concretas (Lógica de Negocio)
        Select Case esp
            Case "Ingeniería en Sistemas"
                cbMateria.Items.AddRange(New String() {"Programación en Visual Basic", "Base de Datos I", "Sistemas Operativos", "Ingeniería de Software", "Redes"})
            Case "Derecho"
                cbMateria.Items.AddRange(New String() {"Derecho Penal", "Derecho Civil", "Lógica Jurídica", "Constitucional"})
            Case "Administración de Empresas"
                cbMateria.Items.AddRange(New String() {"Administración I", "Contabilidad Costos", "Recursos Humanos"})
            Case "Marketing"
                cbMateria.Items.AddRange(New String() {"Mercadotecnia", "Publicidad", "Investigación de Mercados"})
            Case "Diseño Gráfico"
                cbMateria.Items.AddRange(New String() {"Multimedia", "Diseño Editorial", "Ilustración Digital", "Fotografía"})
            Case Else
                cbMateria.Items.Add("General / Pendiente")
        End Select

        If cbMateria.Items.Count > 0 Then cbMateria.SelectedIndex = 0
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        ExportarDataGridViewAExcel(DgvMaterias, "Control de Materias")
    End Sub

    Private Sub btnImprimirReporte_Click(sender As Object, e As EventArgs) Handles btnImprimirReporte.Click
        Try
            If cmbCuatrimestre.Text = "" Then
                MsgBox("Seleccione el Cuatrimestre.", MsgBoxStyle.Exclamation)
                Return
            End If

            Conectar()

            Dim sql As String = ""
            ' Usamos los nombres exactos de tu script: NivelAcademico y Facultad
            If cbDocente.Text = "" Or cbDocente.Text = "--- TODOS LOS DOCENTES ---" Then
                sql = "SELECT D.NombreFull, D.Facultad, D.NivelAcademico, M.NombreMateria, M.Turno, M.Creditos " &
                  "FROM Docentes D " &
                  "INNER JOIN Materias M ON D.IdDocente = M.IdDocente " &
                  "WHERE M.Cuatrimestre = @cuat " &
                  "ORDER BY D.NombreFull ASC"
            Else
                sql = "SELECT D.NombreFull, D.Facultad, D.NivelAcademico, M.NombreMateria, M.Turno, M.Creditos " &
                  "FROM Docentes D " &
                  "INNER JOIN Materias M ON D.IdDocente = M.IdDocente " &
                  "WHERE D.NombreFull = @nom AND M.Cuatrimestre = @cuat"
            End If

            Dim cmd As New SqlCommand(sql, cnn)
            If cbDocente.Text <> "" And cbDocente.Text <> "--- TODOS LOS DOCENTES ---" Then
                cmd.Parameters.AddWithValue("@nom", cbDocente.Text)
            End If
            cmd.Parameters.AddWithValue("@cuat", cmbCuatrimestre.Text)

            Dim da As New SqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            If dt.Rows.Count > 0 Then
                Dim reporte As String = "      AMERICAN COLLEGE - SISTEMA JOSUOS" & vbCrLf &
                                   "      REPORTE ACADÉMICO DE HORARIOS" & vbCrLf &
                                   "      PERIODO: " & cmbCuatrimestre.Text.ToUpper() & vbCrLf &
                                   "==================================================" & vbCrLf

                Dim docenteActual As String = ""

                For Each fila As DataRow In dt.Rows
                    If docenteActual <> fila("NombreFull").ToString() Then
                        docenteActual = fila("NombreFull").ToString()
                        reporte &= vbCrLf & "DOCENTE: " & docenteActual.ToUpper() & vbCrLf &
                               "FACULTAD: " & fila("Facultad").ToString() & vbCrLf &
                               "NIVEL:    " & fila("NivelAcademico").ToString() & vbCrLf & ' Cambiado a NivelAcademico
                               "--------------------------------------------------" & vbCrLf
                    End If

                    reporte &= " * " & fila("NombreMateria").ToString().PadRight(25) &
                           " | " & fila("Turno").ToString() &
                           " | " & fila("Creditos").ToString() & " Cred." & vbCrLf
                Next

                MsgBox(reporte, MsgBoxStyle.Information, "Reporte Generado")
            Else
                MsgBox("No hay materias para este periodo.", MsgBoxStyle.Critical)
            End If

        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        Try
            If cmbCuatrimestre.Text = "" Then
                MsgBox("Seleccione el periodo.") : Return
            End If

            Conectar()
            ' Usamos la consulta de tu script SQL que une Docentes y Materias
            Dim sql As String = ""
            If cbDocente.Text = "" Or cbDocente.Text = "--- TODOS LOS DOCENTES ---" Then
                sql = "SELECT D.NombreFull, D.Facultad, D.NivelAcademico, M.NombreMateria, M.Turno, M.Creditos " &
                  "FROM Docentes D INNER JOIN Materias M ON D.IdDocente = M.IdDocente " &
                  "WHERE M.Cuatrimestre = @cuat ORDER BY D.NombreFull ASC"
            Else
                sql = "SELECT D.NombreFull, D.Facultad, D.NivelAcademico, M.NombreMateria, M.Turno, M.Creditos " &
                  "FROM Docentes D INNER JOIN Materias M ON D.IdDocente = M.IdDocente " &
                  "WHERE D.NombreFull = @nom AND M.Cuatrimestre = @cuat"
            End If

            Dim cmd As New SqlCommand(sql, cnn)
            If cbDocente.Text <> "" And cbDocente.Text <> "--- TODOS LOS DOCENTES ---" Then
                cmd.Parameters.AddWithValue("@nom", cbDocente.Text)
            End If
            cmd.Parameters.AddWithValue("@cuat", cmbCuatrimestre.Text)

            Dim da As New SqlDataAdapter(cmd)
            Dim dtReporte As New DataTable
            da.Fill(dtReporte)

            If dtReporte.Rows.Count > 0 Then
                ' --- AQUÍ GENERAMOS EL PDF ---
                Dim sfd As New SaveFileDialog()
                sfd.Filter = "PDF|*.pdf"
                If sfd.ShowDialog() = DialogResult.OK Then
                    ' Pasamos el DataTable a nuestra función de HTML
                    Dim html As String = ConstruirHTML(dtReporte)

                    Dim tempPath As String = Path.Combine(Path.GetTempPath(), "reporte_josuos.html")
                    File.WriteAllText(tempPath, html)

                    ' Comando de Windows para imprimir a PDF usando Edge (Headless)
                    Dim p As New Process()
                    p.StartInfo.FileName = "msedge.exe"
                    p.StartInfo.Arguments = $"--headless --print-to-pdf=""{sfd.FileName}"" ""{tempPath}"""
                    p.Start()
                    p.WaitForExit()

                    MsgBox("PDF de JosuOs generado correctamente.")
                End If
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub
    Private Function ConstruirHTML(datos As DataTable) As String
        Dim html As String = "<html><head><style>" &
        "body { font-family: 'Segoe UI', Arial; margin: 40px; }" &
        "header { border-bottom: 2px solid #0d47a1; margin-bottom: 20px; }" &
        ".docente-bloque { background: #f9f9f9; padding: 10px; margin-top: 20px; border-left: 5px solid #0d47a1; }" &
        "table { width: 100%; border-collapse: collapse; margin-top: 10px; }" &
        "th { background: #0d47a1; color: white; padding: 10px; }" &
        "td { border: 1px solid #ddd; padding: 8px; }" &
        "</style></head><body>"

        html &= "<header><h1>AMERICAN COLLEGE</h1><h3>REPORTE ACADÉMICO JOSUOS</h3></header>"

        Dim docenteActual As String = ""

        For Each fila As DataRow In datos.Rows
            ' Si cambia el docente, creamos un nuevo encabezado con su Facultad y Nivel
            If docenteActual <> fila("NombreFull").ToString() Then
                If docenteActual <> "" Then html &= "</table>" ' Cerramos tabla anterior

                docenteActual = fila("NombreFull").ToString()
                html &= $"<div class='docente-bloque'>" &
                    $"<b>DOCENTE:</b> {docenteActual.ToUpper()}<br>" &
                    $"<b>FACULTAD:</b> {fila("Facultad")}<br>" &
                    $"<b>NIVEL:</b> {fila("NivelAcademico")}</div>" ' Usando tus nuevas columnas

                html &= "<table><tr><th>Materia</th><th>Turno</th><th>Créditos</th></tr>"
            End If

            html &= $"<tr><td>{fila("NombreMateria")}</td><td>{fila("Turno")}</td><td>{fila("Creditos")}</td></tr>"
        Next

        html &= "</table><p style='margin-top:40px; text-align:center;'>______________________<br>Firma Autorizada</p></body></html>"
        Return html
    End Function
End Class
