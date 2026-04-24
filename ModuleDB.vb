Imports Microsoft.Data.SqlClient
Module ModuleDB

    Public cnn As SqlConnection = New SqlConnection("Data Source=SERVERNAME;Initial Catalog=JosuOsDB;Integrated Security=True;TrustServerCertificate=True")

    Public Sub Conectar()

        Try
            If cnn.State = ConnectionState.Closed Then
                cnn.Open()
            End If

        Catch ex As Exception
            MsgBox("Error al conectar a la base de datos: " & ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    Public Sub Desconectar()

        If cnn.State = ConnectionState.Open Then
            cnn.Close()
        End If

    End Sub

End Module
