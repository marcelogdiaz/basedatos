''' <summary>
''' Clase que implementa la "Fabrica Abstracta" del Patron de Diseño Abstract Factory.
''' Mas info en: https://en.wikipedia.org/wiki/Abstract_factory_pattern
''' </summary>
''' <remarks></remarks>
Public Interface IBD_Factory

    ''' <summary>
    ''' Crea un Visualizador
    ''' </summary>    
    ''' <returns>El Visualizador creado</returns>
    ''' <remarks></remarks>
    Function crear_base_datos() As IBaseDatos

End Interface
