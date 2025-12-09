'*************************************************************************************
'GDGlobal
'               Módulo de declaración de variables globales
'               Converted from VB6 to VB.NET
'*************************************************************************************

Imports System

Namespace Modules

    ' Enumeración de tipos de mensaje
    Public Enum TipoMensaje
        MENSAJE_Informativo = 64 'vbInformation
        MENSAJE_Grave = 16 'vbCritical
        MENSAJE_Exclamacion = 48 'vbExclamation
    End Enum

    ' Tipo de Datos para los Datos del Usuario
    Public Class DatosUsuario
        Public Id As Integer        ' Código del Usuario
        Public Nombre As String     ' Nombre del Usuario
        Public Clave As String      ' Password del Usuario
        Public Instancias As Integer ' Instancias permitidas
        Public NombrePC As String   ' Nombre del PC que puede arrancar (Nothing = todos)
    End Class

    ' Tipo de Datos para la Empresa de Trabajo
    Public Class DatosEmpresaTrabajo
        Public Codigo As Integer
        Public Nombre As String
        Public Servidor As String
        Public BaseDeDatos As String
        Public Usuario As String
        Public Contrasena As String
        Public DSN As String
        Public DLL As String
        Public RutaFicheros As String
        Public RutaFotos As String
        Public RutaInformes As String
        Public Log As String

        Public RutaPaletsCargar As String
        Public RutaPaletsBack As String
        Public RutaPaletsLog As String

        Public RutaVentasCargar As String
        Public RutaVentasBack As String
        Public RutaVentasLog As String

        Public RutaTarifasCargar As String
        Public RutaTarifasBack As String
        Public RutaTarifasLog As String

        Public ServidorRemotoArticulos As String
        Public BaseDeDatosRemotoArticulos As String
        Public UsuarioRemotoArticulos As String
        Public ContrasenaRemotoArticulos As String

        Public CodigoEAN As String
        Public CIF As String
        Public Direccion As String
        Public Poblacion As String
        Public CodigoPostal As String
        Public Pais As Integer

        Public AlmacenLogico As Integer

        Public SoloAsignarDentro As String
        Public SoloAsignarLocal As String
        Public SoloAsignarDeposito As String
        Public SoloAsignarComunitario As String

        Public ContadorTransporte As Integer
    End Class

    ' Tipo de Datos para el Puesto de Trabajo
    Public Class DatosPuestoTrabajo
        Public Id As Integer
        Public Descripcion As String
        Public Corto As String
        Public Impresora As Integer
        Public NombreImpresora As String
        Public NombreAbreviadoImpresora As String
        Public ModeloImpresora As String
        Public LenguajeImpresora As String
        Public EstadoImpresora As Integer
        Public TipoImpresora As String
        Public Colada As Integer
        Public Grupo As Integer
    End Class

    ' Tipo de Datos para las Opciones de Menú
    Public Structure OpcionMenu
        Public Formulario As String
    End Structure

    ' Tipo de Datos para los Menús
    Public Class DatosMenu
        Public Nombre As String
        Public Opcion() As OpcionMenu
    End Class

    Public Module GDGlobal

        '--------------------------------------------------------------------------
        ' Constantes de menú
        Public Const CMD_Aduana As Integer = 0
        Public Const CMD_Almacen As Integer = 1
        Public Const CMD_Compras As Integer = 2
        Public Const CMD_Ventas As Integer = 3
        Public Const CMD_Ficheros As Integer = 4
        Public Const CMD_Estadistica As Integer = 5

        '--------------------------------------------------------------------------
        ' Variables Globales

        ' Variables de la base de datos de configuración
        Public BDDServ As String = ""           ' Nombre del Servidor del Programa Actual
        Public BDDServLocal As String = ""      ' Servidor local
        Public BDDConfig As String = ""         ' Base de datos de Configuración
        Public BDDTime As Integer = 30          ' Tiempo de TimeOut de la Conexión
        Public UsrBDDConfig As String = ""
        Public UsrKeyConfig As String = ""
        Public FicheroDSN As String = ""
        Public FicheroDLL As String = ""
        Public ConexionConfig As String = ""

        ' Variables de trabajo de la empresa
        Public UsrBDD As String = ""
        Public UsrKey As String = ""
        Public CodEmpresa As Integer = 0
        Public ConexionGestion As String = ""
        Public ConexionGestionAlmacen As String = ""
        Public Empresa As String = ""
        Public wRInformes As String = ""
        Public RutaDSN As String = ""

        ' Nombre de la Impresora del Puesto de Trabajo
        Public wImpresora As String = ""

        ' Variable de control de login
        Public LoginSucceeded As Boolean = False

        ' Fichero INI de configuración
        Public ficINI As String = ""

        ' Usuario por defecto
        Public UsrDefault As String = ""

        ' Variable global para la información del usuario
        Public Usuario As DatosUsuario

        ' Variable global para la información de la empresa
        Public EmpresaTrabajo As DatosEmpresaTrabajo

        ' Variable global para el puesto de trabajo
        Public wPuestoTrabajo As DatosPuestoTrabajo

        ' Nombre del PC
        Public nombrePC As String = ""

        ' Variable global para los menús
        Public Menu(5) As DatosMenu

        ' Inicialización del módulo
        Sub New()
            Usuario = New DatosUsuario()
            EmpresaTrabajo = New DatosEmpresaTrabajo()
            wPuestoTrabajo = New DatosPuestoTrabajo()

            For i As Integer = 0 To 5
                Menu(i) = New DatosMenu()
            Next
        End Sub

    End Module

End Namespace
