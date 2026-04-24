Imports Microsoft.Data.SqlClient
Imports System.IO
Public Class UsuariosForm

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles txtConfirmar.TextChanged

    End Sub

    Private Sub UsuariosForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CargarUsuarios()
    End Sub
    Private Sub CargarUsuarios()

        Try

            Conectar()

            Dim query As String = "SELECT ID, Usuario FROM Usuarios"

            Dim da As New SqlDataAdapter(query, cnn)

            Dim tabla As New DataTable()

            da.Fill(tabla)

            DgvUsuarios.DataSource = tabla

            'Cambiar nombres de columnas
            DgvUsuarios.Columns(0).HeaderText = "ID Usuario"
            DgvUsuarios.Columns(1).HeaderText = "Nombre de Usuario"

        Catch ex As Exception

            MsgBox("Error al cargar usuarios: " & ex.Message)

        Finally

            Desconectar()

        End Try

    End Sub

    Private Sub btnSeleccionar_Click(sender As Object, e As EventArgs) Handles btnSeleccionar.Click
        Dim open As New OpenFileDialog()
        open.Filter = "Archivos de imagen|*.jpg;*.jpeg;*.png"

        If open.ShowDialog() = DialogResult.OK Then
            PicUsuario.Image = Image.FromFile(open.FileName)
            PicUsuario.SizeMode = PictureBoxSizeMode.StretchImage
        End If
    End Sub

    Private Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
        'Validar campos
        If txtUsuario.Text = "" Or txtContraseña.Text = "" Then

            MsgBox("Completa todos los campos")

            Return

        End If

        'Validar contraseñas
        If txtContraseña.Text <> txtConfirmar.Text Then

            MsgBox("Las contraseñas no coinciden")

            txtConfirmar.Clear()

            txtConfirmar.Focus()

            Return

        End If

        Try

            Conectar()

            Dim query As String = "INSERT INTO Usuarios (Usuario, Contrasena, Foto) VALUES (@user,@pass,@foto)"

            Dim cmd As New SqlCommand(query, cnn)

            cmd.Parameters.AddWithValue("@user", txtUsuario.Text)

            cmd.Parameters.AddWithValue("@pass", txtContraseña.Text)

            'Guardar imagen
            If PicUsuario.Image IsNot Nothing Then

                Dim ms As New MemoryStream()

                PicUsuario.Image.Save(ms, PicUsuario.Image.RawFormat)

                cmd.Parameters.AddWithValue("@foto", ms.ToArray())

            Else

                cmd.Parameters.AddWithValue("@foto", DBNull.Value)

            End If

            cmd.ExecuteNonQuery()

            MsgBox("Usuario registrado correctamente")

            'Actualizar tabla
            CargarUsuarios()

            'Limpiar campos
            txtUsuario.Clear()
            txtContraseña.Clear()
            txtConfirmar.Clear()
            PicUsuario.Image = Nothing

        Catch ex As Exception

            MsgBox("Error al registrar usuario: " & ex.Message)

        Finally

            Desconectar()

        End Try
    End Sub
End Class
