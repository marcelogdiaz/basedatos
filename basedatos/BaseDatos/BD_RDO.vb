Public Class BD_RDO
    Implements IBaseDatos

    Private conec As New RDO.rdoConnection
    Private dsn As String
    Private mensaje_db As String
    Private resultado As RDO.rdoResultset = Nothing
    Private Const tiempo_espera_sql As Short = 30

    Public Sub Conectar(ByVal valor_dsn As String) Implements IBaseDatos.Conectar
        Try
            dsn = valor_dsn
            conec.Connect = dsn
            conec.LoginTimeout = tiempo_espera_sql
            conec.CursorDriver = MSRDC.CursorDriverConstants.rdUseOdbc
            conec.EstablishConnection()

            mensaje_db = "Se establecio la conexion con " + dsn
            log.Info(mensaje_db)
        Catch ex As Exception
            mensaje_db = "ERROR EN LA CONEXION CON LA BASE DE DATOS"
            MessageBox.Show(mensaje_db)
            log.Error(mensaje_db + " - " + dsn + " - " + ex.Message)
        End Try
    End Sub

    Public Sub Consultar(sql As String) Implements IBaseDatos.Consultar
        'inicializamos la varialble para no retornar valores de otra consulta
        resultado = Nothing

        Try
            resultado = conec.OpenResultset(sql, MSRDC.ResultsetTypeConstants.rdOpenKeyset, MSRDC.LockTypeConstants.rdConcurReadOnly, MSRDC.OptionConstants.rdAsyncEnable + MSRDC.OptionConstants.rdExecDirect)

            While resultado.StillExecuting
            End While

            mensaje_db = "Se realizo la consulta " + sql + " con exito."
            log.Debug(mensaje_db)
        Catch ex As Exception
            mensaje_db = "ERROR AL EJECUTAR LA CONSULTA " + sql
            MessageBox.Show(mensaje_db)
            log.Error(mensaje_db + " - " + dsn + " - " + ex.Message)
        End Try
    End Sub

    Public Sub Desconectar() Implements IBaseDatos.Desconectar
        conec.Close()
        log.Info("Se desconecto " + dsn)
    End Sub

    Public Sub Ejecutar(sql As String) Implements IBaseDatos.Ejecutar
        Try
            conec.Execute(sql)
            bolOperacionCorrecta = True

            mensaje_db = "Se ejecuto la consulta " + sql + " con exito."
            log.Debug(mensaje_db)
        Catch ex As Exception
            ' Capturamos la excepción
            mensaje_db = "ERROR AL EJECUTAR LA CONSULTA " + sql
            MessageBox.Show(mensaje_db)
            log.Error(mensaje_db + " - " + dsn + " - " + ex.Message)
            bolOperacionCorrecta = False
        End Try
    End Sub


    Public Sub Avanzar() Implements IBaseDatos.Avanzar
        resultado.MoveNext()
    End Sub

    Public Function Valor_Columna(columna As String) As Object Implements IBaseDatos.Valor_Columna
        Return IIf(IsDBNull(resultado.rdoColumns(columna).Value), "", resultado.rdoColumns(columna).Value)
    End Function

    Public Function Cantidad_Datos() As Integer Implements IBaseDatos.Cantidad_Datos
        Return resultado.RowCount
    End Function

    Public Function Hay_Datos() As Boolean Implements IBaseDatos.Hay_Datos
        Return Cantidad_Datos() > 0
    End Function
End Class
