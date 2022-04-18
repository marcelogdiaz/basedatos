''' <summary>
''' Clase que implementa una "Fabrica Concreta" del Patron de Diseño Abstract Factory.
''' Esta clase "fabrica" un Visualizador que utiliza GMAP
''' Mas info en: https://en.wikipedia.org/wiki/Abstract_factory_pattern
''' </summary>
''' <remarks></remarks>
Public Class BD_SQL_Factory
    Implements IBD_Factory

    Public Function crear_base_datos() As IBaseDatos Implements IBD_Factory.crear_base_datos
        Return New BD_SQL()
    End Function
End Class
