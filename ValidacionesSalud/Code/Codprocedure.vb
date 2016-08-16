Imports System.Data
Imports MySql.Data.MySqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices


Public Class Codprocedure
    Dim conect As New Conexion

    Public Sub Delete_tablas()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "Delete_tablas"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub
    Public Sub Ceros_izquierda()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_ceros"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub
    Public Sub Hij_papellido()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_HIJPAPELLIDO"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub
    Public Sub Hij_sapellido()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_HIJSAPELLIDO"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub
    Public Sub Hij_snombre()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_HIJSNOMBRE"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub
    Public Sub Sexo_Hombre_Mal()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_sexohombremal"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Public Sub Sexo_Mujer_Mal()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_sexomujermal"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Public Sub Edadnodocumento()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_identificacionerronea"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Public Sub Identificacionrepetida()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_identificacionesrepetidas"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Public Sub Identificacionvacia()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_identificaciovacia"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Public Sub Primerapellidovacio()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_papellidovacio"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Public Sub Pnombrevacio()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_pnombrevacio"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Public Sub Sexonoexiste()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_sexonoexiste"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()

        Conectar_.Close()
    End Sub

    Public Sub Tipodocnoexiste()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_tipodocumentonoexiste"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Public Sub Umenoexiste()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_umenoexiste"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()

    End Sub

    Public Sub Longitudmax()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_logitudmax"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub
    Public Sub Truncatedatoserrores()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_truncatedatosusuerrores"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Public Sub Truncatedatostext()
        Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
        Conectar_.Open()
        Dim cmd As New MySqlCommand
        cmd.Connection = Conectar_
        cmd.CommandText = "PA_truncatedatosusutext"
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()
        Conectar_.Close()
    End Sub

    Private Sub Myobject(ByVal obj As Object)
        Marshal.ReleaseComObject(obj)
        obj = Nothing
    End Sub

    Public Function Llenar() As DataSet
        Try
            Dim myData As New DataSet
            Dim myAdapter As New MySqlDataAdapter
            Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
            Conectar_.Open()
            Dim cmd As New MySqlCommand
            cmd.Connection = Conectar_
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "Estado_Archivo"
            myAdapter.SelectCommand = cmd
            myAdapter.Fill(myData)
            Return myData
        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

    End Function
    Public Function tabla_Errores() As DataTable
        Try
            Dim myData As New DataTable
            Dim myAdapter As New MySqlDataAdapter
            Dim Conectar_ As New MySqlConnection(conect.CrearConexion.ConnectionString)
            Conectar_.Open()
            Dim cmd As New MySqlCommand
            cmd.Connection = Conectar_
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "PA_selectusuerrores"
            myAdapter.SelectCommand = cmd
            myAdapter.Fill(myData)
            Return myData
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function


End Class
