//*************************************************************************************
// Constants.cs
// Módulo para la definición de constantes generales de la aplicación
// Converted from VB.NET to C# for .NET MAUI
//*************************************************************************************

namespace ABG_Almacen_PTL.MAUI.Modules
{
    public static class Constants
    {
        //-------------------------------------------------------------------------------------
        // Constantes para los modos del formulario
        public const int MOD_Seleccion = 0;
        public const int MOD_Edicion = 1;
        public const int MOD_Todo = 2;
        public const int MOD_Nada = 3;

        // Constantes para los botones de acción
        public const int CMD_Salir = 1;
        public const int CMD_Primero = 3;
        public const int CMD_Anterior = 4;
        public const int CMD_Siguiente = 5;
        public const int CMD_Ultimo = 6;
        public const int CMD_Nuevo = 8;
        public const int CMD_Eliminar = 9;
        public const int CMD_Deshacer = 10;
        public const int CMD_Grabar = 11;
        public const int CMD_Pantalla = 13;
        public const int CMD_Imprimir = 14;
        public const int CMD_Filtrar = 16;
        public const int CMD_Buscar = 17;
        public const int CMD_Divisa = 18;
        public const int CMD_Ayuda = 20;
        public const int CMD_Menu = 22;

        public const int MAX_Botones = CMD_Menu;

        // Constantes de Impresion
        public const int CTE_ImpresionPantalla = 0;
        public const int CTE_ImpresionImpresora = 1;
        public const int CTE_CancelarImpresion = 2;

        //-------------------------------------------------------------------------------------
        // Constantes de mensajes
        public const string MSG_001 = " Grabar los cambios? ";
        public const string MSG_002 = " Se regularizarán a 0 las existencias, Continuar? ";
        public const string MSG_003 = " Abandonar los cambios? ";
        public const string MSG_004 = " Imprimir el Formulario? ";
        public const string MSG_005 = " Mensaje Nº 5 ";
        public const string MSG_006 = " No existe: ";
        public const string MSG_007 = " Se Eliminiaran los Datos Permanentemente. Continuar?";

        public const string MSG_050 = " No se ha podido actualizar el artículo! ";
        public const string MSG_051 = " No se ha encontrado el artículo! ";
        public const string MSG_052 = " Grabación Realizada! ";

        // Errores
        public const string MSG_100 = " Error al grabar los datos! ";
        public const string MSG_101 = " Error al borrar los datos! ";
        public const string MSG_102 = " Error, el dato está fuera de rango. ";
        public const string MSG_103 = " Error en el campo: ";
        public const string MSG_104 = " Debe introducir un valor en: ";
        public const string MSG_105 = " se desharán los cambios. ";

        //-------------------------------------------------------------------------------------
        // Constantes de divisas
        public const int DIV_Peseta = 0;     // Divisa Peseta
        public const int DIV_Euro = 1;       // Divisa Euro
        public const int DEC_Peseta = 0;     // Decimales de trabajo en Pesetas
        public const int DEC_Euro = 3;       // Decimales de trabajo en Euros

        // Constantes para la linea de estado
        public const int EST_Mensaje = 1;
        public const int EST_Empresa = 2;
        public const int EST_Divisa = 3;
        public const int EST_Usuario = 4;

        public const int CTE_TiempoEsperaEntornoDatos = 200;
        public const int CTE_TiempoEsperaTransaccion = 10;
        
        // Timeout máximo en segundos para pruebas de conexión inicial
        public const int CTE_TimeoutPruebaConexion = 10;

        public const double MetrosCubicos = 0.028317; // Factor de conversión a m3 de 1 pie3

        // Constantes de Tamaño de Etiquetas para Impresion de Bultos
        public const int ETI_14Con8 = 1;
        public const int ETI_12Con9 = 2;

        // Estados de los Grupos
        public const string EstadoGrupo_Creado = "010";
        public const string EstadoGrupo_Asignado = "020";
        public const string EstadoGrupo_Iniciado = "030";
        public const string EstadoGrupo_Pausado = "040";
        public const string EstadoGrupo_Finalizado = "080";
        public const string EstadoGrupo_Completo = "085";
        public const string EstadoGrupo_Exportado = "090";

        // Constantes SSCC
        public const int IncrementoSerieSSCC_Hipodromo = 30;
    }
}
