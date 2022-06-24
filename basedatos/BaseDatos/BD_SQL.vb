Imports System.Data.SqlClient

Public Class BD_SQL
    Implements IBaseDatos

    Private mensaje_db As String
    Private dsn As String
    Private conec As New SqlConnection()
    Private reader As SqlDataReader
    Private resultados As Collection
    Private puntero_resultados As Integer

    Public Sub Conectar(ByVal valor_dsn As String, Servidor As String, Optional propagar As Boolean = False) Implements IBaseDatos.Conectar

        'Desconectamos automaticamente para evitar errores
        Desconectar()
        Dim cant_puntos As Integer
        Dim parametros As String() = valor_dsn.Split(";")

        Dim catalog As String() = parametros(2).Split("=")

        Try
            dsn = valor_dsn
            'Determinamos si la conexion es LOCAL o REMOTA, esto lo inferimos del valor de "SERVIDOR"
            cant_puntos = Servidor.Split(".").Length - 1

            If (cant_puntos = 3) Then
                'SERVIDOR REMOTO - usamos las credenciales REMOTAS
                conec.ConnectionString = "Server=" + Servidor + ";" + "Initial Catalog=" + catalog(1) + "; User ID=" + userSQLRemoto + ";Password=" + passwordSQLRemoto + "; MultipleActiveResultSets=true;"
            Else
                'SERVIDOR LOCAL
                'conec.ConnectionString = "Data Source=" + Servidor + ";" + "Initial Catalog=" + catalog(1) + "; Integrated Security=true; MultipleActiveResultSets=true;"
                conec.ConnectionString = "Data Source=" + Servidor + ";" + "Initial Catalog=" + catalog(1) + "; User ID=" + userSQLRemoto + ";Password=" + passwordSQLRemoto + "; MultipleActiveResultSets=true;"
            End If

            'Log.Info("DB CONN" + conec.ConnectionString)
            conec.Open()
            mensaje_db = "Se establecio la conexion con " + valor_dsn
            ''log.Info(mensaje_db)
        Catch ex As SqlException
            mensaje_db = "ERROR EN LA CONEXION CON LA BASE DE DATOS: " + Servidor
            'Log.Error("DB CONN" + conec.ConnectionString)
            'Log.Error(mensaje_db + " - " + valor_dsn + " - " + ex.Message)

            If propagar Then
                Throw ex
            Else
                'MessageBox.Show(mensaje_db)
            End If
        End Try
    End Sub

    Public Sub Consultar(sql As String, Optional tipo As CommandType = CommandType.Text, Optional parametros As List(Of SqlParameter) = Nothing) Implements IBaseDatos.Consultar
        Dim SP As SqlCommand

        'inicializamos la varialble para no retornar valores de otra consulta
        resultados = New Collection
        Dim registro As Collection

        puntero_resultados = 1
        Try
            SP = conec.CreateCommand()
            SP.CommandType = tipo
            SP.CommandText = sql

            ''preguntamos si es STORED PROCEDURES, agregamos los parametros
            If (tipo.Equals(CommandType.StoredProcedure)) Then
                'consulta STORED PROCEDURE, recorremos los parametros y agregamos
                SP.Parameters.AddRange(parametros.ToArray())
            End If

            reader = SP.ExecuteReader

            Do While reader.Read()

                Dim registro_actual(reader.FieldCount - 1) As Object
                Dim cantidad_campos As Integer = reader.GetValues(registro_actual)
                registro = New Collection

                'Console.WriteLine("reader.GetValues retrieved {0} columns.", fieldCount)
                For i As Integer = 0 To cantidad_campos - 1
                    registro.Add(New KeyValuePair(Of String, Object)(reader.GetName(i), registro_actual(i)))
                    'Console.Write(values(i).ToString + " - ")
                Next
                resultados.Add(registro)
                'Console.WriteLine()

                'Console.WriteLine("VALOR: {0} - {1} - {2}    *** ", reader.GetValue(0), reader.GetValue(1), reader.GetValue(2))
                'Console.WriteLine("COLUMNA TIPO: {0} ### ", reader.GetOrdinal("Tipo"))
            Loop

            mensaje_db = "Se realizo la consulta " + sql + " con exito."
            ''log.Debug(mensaje_db)

        Catch ex As Exception
            mensaje_db = "ERROR AL EJECUTAR LA CONSULTA " + sql
            'MessageBox.Show(mensaje_db)
            'Log.Error(mensaje_db + " - " + dsn + " - " + ex.Message)
        End Try
    End Sub

    Public Sub Desconectar() Implements IBaseDatos.Desconectar
        conec.Dispose()
        'conec.Close()
        ''log.Info("Se desconecto " + dsn)
    End Sub

    Public Sub Ejecutar(sql As String) Implements IBaseDatos.Ejecutar
        Dim SP As SqlCommand
        Try
            SP = conec.CreateCommand()
            SP.CommandText = sql
            reader = SP.ExecuteReader()

            mensaje_db = "Se ejecuto la consulta " + sql + " con exito."
            ''log.Debug(mensaje_db)
        Catch ex As Exception
            ' Capturamos la excepción
            mensaje_db = "ERROR AL EJECUTAR LA CONSULTA " + sql
            'MessageBox.Show(mensaje_db)
            'Log.Error(mensaje_db + " - " + dsn + " - " + ex.Message)
            Throw ex
        End Try
    End Sub


    Public Sub Avanzar() Implements IBaseDatos.Avanzar
        puntero_resultados = puntero_resultados + 1
    End Sub

    Public Function Valor_Columna(columna As String) As Object Implements IBaseDatos.Valor_Columna
        'recorrer RESULTADOS y buscar valor de llave COLUMNA
        Dim registro As Collection = CType(resultados.Item(puntero_resultados), Collection)
        For Each item As KeyValuePair(Of String, Object) In registro
            If (String.Compare(item.Key, columna, True) = 0) Then 'If (item.Key.Equals(columna)) Then
                'IMPORTANTE: si el valor del cambo en la Base de Datos es NULL => se asigna "" 
                'para evitar una propagacion de ERRORES
                Return IIf(IsDBNull(item.Value), "", item.Value)
            End If
        Next

        Return Nothing
    End Function

    Public Function Cantidad_Datos() As Integer Implements IBaseDatos.Cantidad_Datos
        Return resultados.Count
    End Function

    Public Function Hay_Datos() As Boolean Implements IBaseDatos.Hay_Datos
        Return Cantidad_Datos() > 0
    End Function

    Public Function Registros() As Collection Implements IBaseDatos.Registros
        Return resultados
    End Function
End Class
