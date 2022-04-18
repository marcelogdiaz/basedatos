Module GestorBD

    Private BaseDatosF As IBD_Factory
    Private bd As IBaseDatos

    Private NombreBD As String
    Public miNombreEquipo As String = Environment.MachineName
    Public instanciaSQL As String = "\SQLEXPRESS"
    ''' <summary>
    ''' Nombre del "Data Source" a usar en la conexion a BD SAR. Por defecto es local a la PC, puede cambiarse mediante entrada en @archivoINI
    ''' </summary>
    Public ServidorSQL As String = miNombreEquipo + instanciaSQL

    ''credenciales POR DEFECTO 
    Public userSQLRemoto As String = "SA"
    Public passwordSQLRemoto As String = ""

    Public Sub inicializarGestorBD(nombreBaseDatos As String, ByVal user As String, ByVal pwd As String)
        'BaseDatosF = New BD_RDO_Factory
        BaseDatosF = New BD_SQL_Factory
        bd = BaseDatosF.crear_base_datos()
        NombreBD = nombreBaseDatos

        ''actualizamos las credenciales de conexion
        userSQLRemoto = user
        passwordSQLRemoto = pwd
    End Sub

    Public Function ejecutar(texto_sql As String) As Collection
        Dim resultados As New Collection
        bd.Conectar("UID=" + userSQLRemoto + ";PWD=" + passwordSQLRemoto + ";DSN=" + NombreBD, ServidorSQL)

        bd.Consultar(texto_sql)

        If bd.Hay_Datos() Then
            resultados = bd.Registros
        End If
        bd.Desconectar()

        Return resultados
    End Function

    ''' <summary>
    ''' Funcion wrapper para hacer un SELECT al motor de BD
    ''' </summary>
    ''' <param name="tabla">String con nombre de la tabla de la BD</param>
    ''' <param name="where">String OPCIONAL, que debe tener toda la condicion del WHERE</param>
    ''' <returns>Collection con resultados de la consulta BD</returns>
    Public Function selectar(tabla As String, Optional ByVal where As String = "") As Collection
        Dim resultados As New Collection
        Dim texto_sql As String = ""
        bd.Conectar("UID=" + userSQLRemoto + ";PWD=" + passwordSQLRemoto + ";DSN=" + NombreBD, ServidorSQL)

        'BUSCAMOS SOLO CASOS ABIERTOS
        texto_sql = "SELECT * FROM " + tabla

        If (Not where.Equals("")) Then
            texto_sql = texto_sql + " where " + where
        End If

        bd.Consultar(texto_sql)

        If bd.Hay_Datos() Then
            resultados = bd.Registros
        End If
        bd.Desconectar()

        Return resultados
    End Function

    ''' <summary>
    ''' Funcion wrapper para hacer un SELECT al motor de BD
    ''' </summary>
    ''' <param name="tabla">String con nombre de la tabla de la BD</param>
    ''' <param name="where">String OPCIONAL, que debe tener toda la condicion del WHERE</param>
    ''' <returns>Collection con resultados de la consulta BD</returns>
    Public Function selectar_un_registro(tabla As String, Optional ByVal where As String = "") As Collection
        Dim resultados As New Collection
        Dim texto_sql As String = ""
        bd.Conectar("UID=" + userSQLRemoto + ";PWD=" + passwordSQLRemoto + ";DSN=" + NombreBD, ServidorSQL)

        'BUSCAMOS SOLO CASOS ABIERTOS
        texto_sql = "SELECT * FROM " + tabla

        If (Not where.Equals("")) Then
            texto_sql = texto_sql + " where " + where
        End If

        bd.Consultar(texto_sql)

        If bd.Hay_Datos() Then
            resultados = bd.Registros
            Return resultados.Item(1) '' ojo ver
        End If
        bd.Desconectar()
        Return resultados
    End Function

    ''' <summary>
    ''' Funcion wrapper para hacer un UPDATE al motor de BD
    ''' </summary>
    ''' <param name="tabla">String de la tabla BD a actualizar</param>
    ''' <param name="campos">Collection de KeyValuePair(NOMBRECAMPO, VALORCAMPO)</param>
    ''' <param name="where">String OPCIONAL, que debe tener toda la condicion del WHERE</param>
    Public Sub actualizar(tabla As String, campos As Collection, Optional ByVal where As String = "")
        Dim texto_sql As String = ""
        bd.Conectar("UID=" + userSQLRemoto + ";PWD=" + passwordSQLRemoto + ";DSN=" + NombreBD, ServidorSQL)

        texto_sql = "UPDATE " + tabla + " SET "
        For Each pair As KeyValuePair(Of String, String) In campos
            texto_sql = texto_sql + pair.Key + "=" + Trim(pair.Value) + ","
        Next

        'Le sacamos la ultima "COMA" antes del WHERE
        texto_sql = texto_sql.Substring(0, texto_sql.Length() - 1)

        If (Not where.Equals("")) Then
            texto_sql = texto_sql + " where " + where
        End If
        Try
            bd.Ejecutar(texto_sql)
        Catch ex As Exception
            'log.Error("GestorDB.actualizar - " + ex.Message)
        Finally
            bd.Desconectar()
        End Try
    End Sub

    ''' <summary>
    ''' Funcion wrapper para hacer un INSERT en la BD
    ''' </summary>
    ''' <param name="tabla">String de la tabla BD a insertar</param>
    ''' <param name="campos">Collection de KeyValuePair(NOMBRECAMPO, VALORCAMPO)</param>
    ''' <param name="where">String OPCIONAL, que debe tener toda la condicion del WHERE</param>
    Public Function insertar(tabla As String, campos As Collection, Optional ByVal where As String = "") As Integer
        Dim valorID As Integer
        Dim texto_sql As String = ""
        bd.Conectar("UID=" + userSQLRemoto + ";PWD=" + passwordSQLRemoto + ";DSN=" + NombreBD, ServidorSQL)

        texto_sql = "INSERT INTO " + tabla + " ("
        'colocamos los nombre de los campos a insetar

        For Each pair As KeyValuePair(Of String, String) In campos
            texto_sql = texto_sql + Trim(pair.Key) + ","
        Next
        'Le sacamos la ultima "COMA" 
        texto_sql = texto_sql.Substring(0, texto_sql.Length() - 1)
        texto_sql = texto_sql + ") VALUES ("

        'colocamos los valores de los campos a insetar
        For Each pair As KeyValuePair(Of String, String) In campos
            texto_sql = texto_sql + pair.Value + ","
        Next

        'Le sacamos la ultima "COMA" 
        texto_sql = texto_sql.Substring(0, texto_sql.Length() - 1)
        texto_sql = texto_sql + ") "

        If (Not where.Equals("")) Then
            texto_sql = texto_sql + " where " + where
        End If

        Try
            bd.Ejecutar(texto_sql)

            'retornamos el ID insertado siempre y cuando SEA UN AUTOINCREMENT, SINO retornamos 0
            bd.Consultar("select SCOPE_IDENTITY() as [ultimoid]")
            If (bd.Valor_Columna("ultimoid").Equals("")) Then
                valorID = 0
            Else
                valorID = CInt(bd.Valor_Columna("ultimoid"))
            End If
        Catch ex As Exception
            'log.Error("GestorDB.insertar - " + ex.Message)
        Finally
            bd.Desconectar()
        End Try
        Return valorID
    End Function

    ''' <summary>
    ''' Funcion wrapper para hacer un DELETE en la BD
    ''' </summary>
    ''' <param name="tabla">String de la tabla BD a insertar</param>
    ''' <param name="where">String OPCIONAL, que debe tener toda la condicion del WHERE</param>
    Public Sub borrar(tabla As String, Optional ByVal where As String = "")
        Dim texto_sql As String = ""
        bd.Conectar("UID=" + userSQLRemoto + ";PWD=" + passwordSQLRemoto + ";DSN=" + NombreBD, ServidorSQL)

        texto_sql = "DELETE FROM " + tabla
        If (Not where.Equals("")) Then
            texto_sql = texto_sql + " where " + where
        End If

        Try
            bd.Ejecutar(texto_sql)
        Catch ex As Exception
            'log.Error("GestorDB.borrar - " + ex.Message)
        Finally
            bd.Desconectar()
        End Try
    End Sub

    Public Function Valor_Columna(columna As String, valores As Collection) As Object
        'recorrer RESULTADOS y buscar valor de llave COLUMNA
        'Dim registro As Collection = CType(resultados.Item(puntero_resultados), Collection)
        For Each item As KeyValuePair(Of String, Object) In valores
            If (String.Compare(item.Key, columna, True) = 0) Then 'If (item.Key.Equals(columna)) Then
                'IMPORTANTE: si el valor del cambo en la Base de Datos es NULL => se asigna "" 
                'para evitar una propagacion de ERRORES
                Return IIf(IsDBNull(item.Value), "", item.Value)
            End If
        Next

        Return Nothing
    End Function

End Module
