Imports Microsoft.Data.SqlClient
Imports System.Data
Public Class DocentesForm
    Private Sub ListarDocentes()
        Try
            Conectar()
            Dim da As New SqlDataAdapter("SELECT * FROM Docentes", cnn)
            Dim dt As New DataTable
            da.Fill(dt)
            DgvDocentes.DataSource = dt
        Catch ex As Exception
            MsgBox("Error al listar: " & ex.Message)
        Finally
            cnn.Close()
        End Try
    End Sub
    Private Sub DocentesForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Limpiamos por si acaso
        cmbEspecialidad.Items.Clear()

        ' Agregamos las carreras de la American College
        cmbEspecialidad.Items.Add("Ingeniería en Sistemas")
        cmbEspecialidad.Items.Add("Derecho")
        cmbEspecialidad.Items.Add("Administración de Empresas")
        cmbEspecialidad.Items.Add("Marketing")
        cmbEspecialidad.Items.Add("Diseño Gráfico")

        ' Seleccionar el primero por defecto
        cmbEspecialidad.SelectedIndex = 0

        ' También cargamos la tabla del Grid
        ListarDocentes()
        ' Agregamos los grados académicos
        cmbNivel.Items.Clear()
        cmbNivel.Items.Add("Licenciatura")
        cmbNivel.Items.Add("Ingeniería")
        cmbNivel.Items.Add("Postgrado")
        cmbNivel.Items.Add("Maestría")
        cmbNivel.Items.Add("Doctorado")

        ' Seleccionar Maestría por defecto (es común en la UAC)
        cmbNivel.SelectedIndex = 3
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Try
            Conectar()

            ' Query actualizado con las nuevas columnas
            Dim query As String = "INSERT INTO Docentes (NombreFull, Cedula, Especialidad, Telefono, Email, NivelAcademico, Facultad) " &
                                   "VALUES (@nom, @ced, @esp, @tel, @mail, @nivel, @fac)"

            Dim cmd As New SqlCommand(query, cnn)

            cmd.Parameters.AddWithValue("@nom", txtNombre.Text)
            cmd.Parameters.AddWithValue("@ced", txtCedula.Text)
            cmd.Parameters.AddWithValue("@esp", cmbEspecialidad.Text)
            cmd.Parameters.AddWithValue("@tel", txtTelefono.Text)
            cmd.Parameters.AddWithValue("@mail", txtEmail.Text)
            ' Nuevos parámetros
            cmd.Parameters.AddWithValue("@nivel", cmbNivel.Text)
            cmd.Parameters.AddWithValue("@fac", txtFacultad.Text)

            cmd.ExecuteNonQuery()
            MsgBox("Docente guardado correctamente", MsgBoxStyle.Information)

            ListarDocentes()
            LimpiarCampos() ' Es buena práctica limpiar los controles después de guardar

        Catch ex As Exception
            MsgBox("Error al guardar: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            Desconectar()
        End Try
    End Sub
    Private Sub LimpiarCampos()
        txtNombre.Clear()
        txtCedula.Clear()
        txtTelefono.Clear()
        txtEmail.Clear()
        txtFacultad.Clear()
        cmbEspecialidad.SelectedIndex = 0
        If cmbNivel.Items.Count > 0 Then cmbNivel.SelectedIndex = 0
    End Sub

    Private Sub btnElminar_Click(sender As Object, e As EventArgs) Handles btnElminar.Click
        Try
            ' 1. Validar selección
            If DgvDocentes.CurrentRow Is Nothing Then
                MsgBox("Seleccione una fila primero")
                Return
            End If

            ' 2. Capturar el ID (Usando el índice 0 para asegurar)
            ' Más seguro que usar el índice 0
            Dim idDoc As String = DgvDocentes.CurrentRow.Cells("IdDocente").Value.ToString()

            ' MENSAJE DE DEPURACIÓN (Bórralo cuando funcione)
            ' MsgBox("Intentando borrar el ID: " & idDoc)

            Dim respuesta As DialogResult = MessageBox.Show("¿Borrar docente con ID " & idDoc & "?", "Alerta", MessageBoxButtons.YesNo)

            If respuesta = DialogResult.Yes Then
                Conectar()
                Dim sql As String = "DELETE FROM Docentes WHERE IdDocente = @id"
                Dim cmd As New SqlCommand(sql, cnn)
                cmd.Parameters.AddWithValue("@id", idDoc)

                Dim filasAfectadas As Integer = cmd.ExecuteNonQuery()

                If filasAfectadas > 0 Then
                    MsgBox("Eliminado con éxito")
                Else
                    MsgBox("No se encontró el registro en la base de datos")
                End If

                ListarDocentes() ' RECARGAR EL GRID
            End If

        Catch ex As Exception
            MsgBox("Error crítico: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        ExportarDataGridViewAExcel(DgvDocentes, "Control de Docentes")
    End Sub
End Class
