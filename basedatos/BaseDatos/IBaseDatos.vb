
Imports System.Data.SqlClient

Public Interface IBaseDatos

    ''' <summary>
    ''' Permite conectarse con una fuente de datos
    ''' </summary>
    ''' <param name="valor_dsn">Parametros de conexion</param>
    ''' <remarks></remarks>
    Sub Conectar(ByVal valor_dsn As String, Servidor As String, Optional propagar As Boolean = False)

    ''' <summary>
    ''' Permite realizar una consulta SQL a la fuente de datos y almacena el resultado en una estructura interna
    ''' </summary>
    ''' <param name="sql">Consulta SQL</param>
    ''' <remarks></remarks>
    Sub Consultar(ByVal sql As String, Optional tipo As CommandType = CommandType.Text, Optional parametros As List(Of SqlParameter) = Nothing, Optional propagar As Boolean = False)

    ''' <summary>
    ''' Permite desconectarse de la fuente de datos
    ''' </summary>
    ''' <remarks></remarks>
    Sub Desconectar()

    ''' <summary>
    ''' Permite realizar una consulta SQL a la fuente de datos.
    ''' </summary>
    ''' <param name="sql">Consulta SQL</param>
    ''' <remarks></remarks>
    Sub Ejecutar(ByVal sql As String, Optional propagar As Boolean = False)

    ''' <summary>
    ''' Permite avanzar una posicion del cursor de la estructura interna de almacenamiento
    ''' </summary>
    ''' <remarks></remarks>
    Sub Avanzar()

    ''' <summary>
    ''' Permite obtener el valor de la columna del registro actual dentro de la estructura interna de almacenamiento
    ''' </summary>
    ''' <param name="columna">Nombre de la columna</param>
    ''' <returns>Valor de la columna o NOTHING si no se encuentra la columna</returns>
    ''' <remarks></remarks>
    Function Valor_Columna(ByVal columna As String) As Object

    ''' <summary>
    ''' Permite obtener cuantos elementos contiene la estructura interna de almacenamiento
    ''' </summary>
    ''' <returns>Cantidad de datos de la estructura interna de almacenamiento</returns>
    ''' <remarks></remarks>
    Function Cantidad_Datos() As Integer

    ''' <summary>
    ''' Permite consultar si la estructura interna de almacenamiento tiene datos
    ''' </summary>
    ''' <returns>Si/No</returns>
    ''' <remarks></remarks>
    Function Hay_Datos() As Boolean

    Function Registros() As Collection

End Interface
