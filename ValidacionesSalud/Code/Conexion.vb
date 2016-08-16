Imports MySql.Data.MySqlClient
Public Class Conexion

    Public Function CrearConexion() As MySqlConnection
        Try
            Dim dato_Con As String = "Data Source=localhost; Database=validaciones; User Id=root; Password=''"
            Dim MySqlConexion As New MySqlConnection(dato_Con)
            Return MySqlConexion
        Catch ex As Exception
            MsgBox("Error de conexion", MsgBoxStyle.Information, "Validacion")
        End Try
    End Function



End Class
