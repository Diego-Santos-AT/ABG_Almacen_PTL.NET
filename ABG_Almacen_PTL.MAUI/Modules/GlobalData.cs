//*************************************************************************************
// GlobalData.cs
// Módulo de declaración de variables globales
// Converted from VB.NET to C# for .NET MAUI
//*************************************************************************************

namespace ABG_Almacen_PTL.MAUI.Modules
{
    // Enumeración de tipos de mensaje
    public enum TipoMensaje
    {
        MENSAJE_Informativo = 64,    // vbInformation
        MENSAJE_Grave = 16,          // vbCritical
        MENSAJE_Exclamacion = 48     // vbExclamation
    }

    // Tipo de Datos para los Datos del Usuario
    public class DatosUsuario
    {
        public int Id { get; set; }              // Código del Usuario
        public string Nombre { get; set; } = "";  // Nombre del Usuario
        public string Clave { get; set; } = "";   // Password del Usuario
        public int Instancias { get; set; }       // Instancias permitidas
        public string? NombrePC { get; set; }     // Nombre del PC que puede arrancar (null = todos)
    }

    // Tipo de Datos para la Empresa de Trabajo
    public class DatosEmpresaTrabajo
    {
        public int Codigo { get; set; }
        public string Nombre { get; set; } = "";
        public string Servidor { get; set; } = "";
        public string BaseDeDatos { get; set; } = "";
        public string Usuario { get; set; } = "";
        public string Contrasena { get; set; } = "";
        public string DSN { get; set; } = "";
        public string DLL { get; set; } = "";
        public string RutaFicheros { get; set; } = "";
        public string RutaFotos { get; set; } = "";
        public string RutaInformes { get; set; } = "";
        public string Log { get; set; } = "";

        public string RutaPaletsCargar { get; set; } = "";
        public string RutaPaletsBack { get; set; } = "";
        public string RutaPaletsLog { get; set; } = "";

        public string RutaVentasCargar { get; set; } = "";
        public string RutaVentasBack { get; set; } = "";
        public string RutaVentasLog { get; set; } = "";

        public string RutaTarifasCargar { get; set; } = "";
        public string RutaTarifasBack { get; set; } = "";
        public string RutaTarifasLog { get; set; } = "";

        public string ServidorRemotoArticulos { get; set; } = "";
        public string BaseDeDatosRemotoArticulos { get; set; } = "";
        public string UsuarioRemotoArticulos { get; set; } = "";
        public string ContrasenaRemotoArticulos { get; set; } = "";

        public string CodigoEAN { get; set; } = "";
        public string CIF { get; set; } = "";
        public string Direccion { get; set; } = "";
        public string Poblacion { get; set; } = "";
        public string CodigoPostal { get; set; } = "";
        public int Pais { get; set; }

        public int AlmacenLogico { get; set; }

        public string SoloAsignarDentro { get; set; } = "";
        public string SoloAsignarLocal { get; set; } = "";
        public string SoloAsignarDeposito { get; set; } = "";
        public string SoloAsignarComunitario { get; set; } = "";

        public int ContadorTransporte { get; set; }
    }

    // Tipo de Datos para el Puesto de Trabajo
    public class DatosPuestoTrabajo
    {
        public int Id { get; set; }
        public string Descripcion { get; set; } = "";
        public string Corto { get; set; } = "";
        public int Impresora { get; set; }
        public string NombreImpresora { get; set; } = "";
        public string NombreAbreviadoImpresora { get; set; } = "";
        public string ModeloImpresora { get; set; } = "";
        public string LenguajeImpresora { get; set; } = "";
        public int EstadoImpresora { get; set; }
        public string TipoImpresora { get; set; } = "";
        public int Colada { get; set; }
        public int Grupo { get; set; }
    }

    // Tipo de Datos para las Opciones de Menú
    public struct OpcionMenu
    {
        public string Formulario { get; set; }
    }

    // Tipo de Datos para los Menús
    public class DatosMenu
    {
        public string Nombre { get; set; } = "";
        public OpcionMenu[]? Opcion { get; set; }
    }

    // Variables globales de la aplicación
    public static class GlobalData
    {
        //--------------------------------------------------------------------------
        // Constantes de menú
        public const int CMD_Aduana = 0;
        public const int CMD_Almacen = 1;
        public const int CMD_Compras = 2;
        public const int CMD_Ventas = 3;
        public const int CMD_Ficheros = 4;
        public const int CMD_Estadistica = 5;

        //--------------------------------------------------------------------------
        // Variables Globales

        // Variables de la base de datos de configuración
        public static string BDDServ { get; set; } = "";           // Nombre del Servidor del Programa Actual
        public static string BDDServLocal { get; set; } = "";      // Servidor local
        public static string BDDConfig { get; set; } = "";         // Base de datos de Configuración
        public static int BDDTime { get; set; } = 30;              // Tiempo de TimeOut de la Conexión
        public static string UsrBDDConfig { get; set; } = "";
        public static string UsrKeyConfig { get; set; } = "";
        public static string FicheroDSN { get; set; } = "";
        public static string FicheroDLL { get; set; } = "";
        public static string ConexionConfig { get; set; } = "";

        // Variables de trabajo de la empresa
        public static string UsrBDD { get; set; } = "";
        public static string UsrKey { get; set; } = "";
        public static int CodEmpresa { get; set; } = 0;
        public static string ConexionGestion { get; set; } = "";
        public static string ConexionGestionAlmacen { get; set; } = "";
        public static string Empresa { get; set; } = "";
        public static string wRInformes { get; set; } = "";
        public static string RutaDSN { get; set; } = "";

        // Nombre de la Impresora del Puesto de Trabajo
        public static string wImpresora { get; set; } = "";

        // Variable de control de login
        public static bool LoginSucceeded { get; set; } = false;

        // Fichero INI de configuración (replaced with Preferences in MAUI)
        public static string ficINI { get; set; } = "";

        // Usuario por defecto
        public static string UsrDefault { get; set; } = "";

        // Variable global para la información del usuario
        public static DatosUsuario Usuario { get; set; } = new DatosUsuario();

        // Variable global para la información de la empresa
        public static DatosEmpresaTrabajo EmpresaTrabajo { get; set; } = new DatosEmpresaTrabajo();

        // Variable global para el puesto de trabajo
        public static DatosPuestoTrabajo wPuestoTrabajo { get; set; } = new DatosPuestoTrabajo();

        // Nombre del PC
        public static string nombrePC { get; set; } = "";

        // Variable global para los menús
        public static DatosMenu[] Menu { get; set; } = new DatosMenu[6];

        // Inicialización estática
        static GlobalData()
        {
            Usuario = new DatosUsuario();
            EmpresaTrabajo = new DatosEmpresaTrabajo();
            wPuestoTrabajo = new DatosPuestoTrabajo();

            for (int i = 0; i <= 5; i++)
            {
                Menu[i] = new DatosMenu();
            }
        }
    }
}
