Imports Microsoft.Data.SqlClient
Imports System.IO
Public Class LoginForm

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Try
            Conectar()

            Dim query As String = "SELECT Usuario, Foto FROM Usuarios WHERE Usuario = @user AND Contrasena = @pass"
            Dim cmd As New SqlCommand(query, cnn)

            cmd.Parameters.AddWithValue("@user", txtUsuario.Text)
            cmd.Parameters.AddWithValue("@pass", txtContraseña.Text)

            Dim lector As SqlDataReader = cmd.ExecuteReader()

            If lector.Read() Then
                ' Cargar la foto de perfil del usuario en el Dashboard
                If Not IsDBNull(lector("Foto")) Then
                    Dim imgBytes As Byte() = DirectCast(lector("Foto"), Byte())
                    Dim ms As New MemoryStream(imgBytes)
                    DeskopForm.picFotoPerfil.Image = Image.FromStream(ms)
                End If

                MsgBox("Bienvenido a ACN LINUX, " & lector("Usuario").ToString(), MsgBoxStyle.Information)

                DeskopForm.Show()
                Me.Hide()
            Else
                MsgBox("Usuario o contraseña incorrectos.", MsgBoxStyle.Critical)
            End If

        Catch ex As Exception
            MsgBox("Error de conexión: " & ex.Message)
        Finally
            Desconectar()
        End Try
    End Sub

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Configuración inicial: La contraseña inicia oculta
        txtContraseña.PasswordChar = "*"c
        rbOcultar.Checked = True
    End Sub

    Private Sub rbMostrar_CheckedChanged(sender As Object, e As EventArgs) Handles rbMostrar.CheckedChanged
        If rbMostrar.Checked Then
            ' Muestra el texto tal cual
            txtContraseña.PasswordChar = ControlChars.NullChar
        End If
    End Sub

    Private Sub rbOcultar_CheckedChanged(sender As Object, e As EventArgs) Handles rbOcultar.CheckedChanged
        If rbOcultar.Checked Then
            ' Oculta el texto con asteriscos
            txtContraseña.PasswordChar = "*"c
        End If
    End Sub
End Class
