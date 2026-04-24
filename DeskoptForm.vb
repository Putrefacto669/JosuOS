Public Class DeskopForm
    Private Sub btnInicio_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DeskopForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Asegurarnos de que el panel inicie oculto
        Panel1.Visible = False
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        ' 1. Alternar visibilidad
        Panel1.Visible = Not Panel1.Visible

        If Panel1.Visible Then
            ' 2. Forzar el tamaño antes de calcular la posición
            Panel1.Size = New Size(200, 300)

            ' 3. TRUCO MAESTRO: Corregido a PictureBox1 (con el número)
            Dim puntoEnPantalla As Point = PictureBox1.PointToScreen(Point.Empty)
            Dim puntoEnFormulario As Point = Me.PointToClient(puntoEnPantalla)

            ' 4. Posicionar: Misma X que el PictureBox, pero en Y subimos el alto del panel
            Dim posX As Integer = puntoEnFormulario.X
            Dim posY As Integer = puntoEnFormulario.Y - Panel1.Height

            Panel1.Location = New Point(posX, posY)

            ' 5. Asegurar que esté por encima de todo
            Panel1.BringToFront()
        End If
    End Sub

    Private Sub PicCalculadora_Click(sender As Object, e As EventArgs) Handles PicCalculadora.Click
        ' 1. Creamos la instancia del formulario
        Dim fCalc As New CalculadoraForm()

        ' 2. Lo configuramos para que no se salga de nuestro DesktopForm
        fCalc.TopLevel = False
        Me.Controls.Add(fCalc) ' Se agrega al escritorio

        ' 3. Lo mostramos y lo traemos al frente
        fCalc.Show()
        fCalc.BringToFront()

        ' 4. Cerramos el menú de inicio
        Panel1.Visible = False
    End Sub

    Private Sub PicUsuarios_Click(sender As Object, e As EventArgs) Handles PicUsuarios.Click
        Dim fUsuarios As New UsuariosForm()
        fUsuarios.TopLevel = False
        Me.Controls.Add(fUsuarios)
        fUsuarios.Show()
        fUsuarios.BringToFront()
        Panel1.Visible = False
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Dim FDocentes As New DocentesForm()
        FDocentes.TopLevel = False
        Me.Controls.Add(FDocentes)
        FDocentes.Show()
        FDocentes.BringToFront()
        Panel1.Visible = False
    End Sub

    Private Sub flpBarraTareas_Paint(sender As Object, e As PaintEventArgs) Handles flpBarraTareas.Paint

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Dim respuesta As DialogResult = MessageBox.Show("¿Desea cerrar ACN LINUX y salir del sistema?",
                                                        "Cerrar Sistema",
                                                        MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question)
        If respuesta = DialogResult.Yes Then
            Application.Exit() ' Esto cierra todo el programa de un solo
        End If
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Dim fMaterias As New FormMaterias()
        fMaterias.TopLevel = False
        Me.Controls.Add(fMaterias)
        fMaterias.Show()
        fMaterias.BringToFront()
        Panel1.Visible = False
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        Dim fAsistencias As New FormAsistencias()
        fAsistencias.TopLevel = False
        Me.Controls.Add(fAsistencias)
        fAsistencias.Show()
        fAsistencias.BringToFront()
        Panel1.Visible = False
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        Dim fReportes As New FormExamenes()
        fReportes.TopLevel = False
        Me.Controls.Add(fReportes)
        fReportes.Show()
        fReportes.BringToFront()
        Panel1.Visible = False
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        ' 1. DETECTAR FORMULARIO ACTIVO
        Dim frmAsis As FormAsistencias = Application.OpenForms.OfType(Of FormAsistencias)().FirstOrDefault()
        Dim frmExam As FormExamenes = Application.OpenForms.OfType(Of FormExamenes)().FirstOrDefault()
        Dim frmMat As FormMaterias = Application.OpenForms.OfType(Of FormMaterias)().FirstOrDefault()

        Dim dgvSeleccionado As DataGridView = Nothing
        Dim nombreReporte As String = ""

        If frmAsis IsNot Nothing AndAlso frmAsis.Visible Then
            dgvSeleccionado = frmAsis.dgvAsistencias
            nombreReporte = "Reporte de Asistencias"
        ElseIf frmExam IsNot Nothing AndAlso frmExam.Visible Then
            dgvSeleccionado = frmExam.dgvExamenes
            nombreReporte = "Constancia de Exámenes"
        ElseIf frmMat IsNot Nothing AndAlso frmMat.Visible Then
            dgvSeleccionado = frmMat.DgvMaterias
            nombreReporte = "Lista de Materias"
        End If

        If dgvSeleccionado Is Nothing Then
            MsgBox("Por favor, abra un formulario con datos primero.", MsgBoxStyle.Exclamation, "JosuOs")
            Return
        End If

        ' 2. MENÚ DE OPCIONES
        Dim rpta = MsgBox("¿Qué desea hacer?" & vbCrLf & vbCrLf &
                      "[SÍ] - Solo generar y guardar PDF" & vbCrLf &
                      "[NO] - Enviar un archivo por Correo",
                      MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, "Centro de Reportes JosuOs")

        Select Case rpta
            Case MsgBoxResult.Yes
                ' Solo generamos el PDF (Windows pedirá dónde guardarlo)
                ExportarAPDFGlobal(dgvSeleccionado, nombreReporte)

            Case MsgBoxResult.No
                ' El usuario primero debe tener el PDF guardado
                MsgBox("Seleccione el archivo PDF que desea enviar.", MsgBoxStyle.Information, "Adjuntar Reporte")

                Dim OFD As New OpenFileDialog()
                OFD.Filter = "Archivos PDF (*.pdf)|*.pdf"
                OFD.Title = "Seleccione el reporte de " & nombreReporte

                If OFD.ShowDialog() = DialogResult.OK Then
                    Dim correoDestino As String = InputBox("Ingrese el correo del destinatario:", "Enviar " & nombreReporte)

                    If Not String.IsNullOrEmpty(correoDestino) Then
                        ' Llamamos al módulo de Email con la ruta que el usuario eligió
                        EnviarReportePorEmail(correoDestino, OFD.FileName, "Reporte Académico: " & nombreReporte)
                    End If
                End If

            Case MsgBoxResult.Cancel
                Return
        End Select
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub MenuStrip2_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub

    Private Sub picFotoPerfil_DoubleClick(sender As Object, e As EventArgs) Handles picFotoPerfil.DoubleClick
        ' ¡Sorpresa!
        Dim egg As New FormPaint()
        egg.Show()
    End Sub
End Class
