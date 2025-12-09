'*************************************************************************************
'GDConstantes
'               Módulo para la definición de constantes generales de la aplicación
'               Converted from VB6 to VB.NET
'*************************************************************************************

Imports System

Namespace Modules

    Public Module GDConstantes

        '-------------------------------------------------------------------------------------
        'Constantes para los modos del formulario
        Public Const MOD_Seleccion As Integer = 0
        Public Const MOD_Edicion As Integer = 1
        Public Const MOD_Todo As Integer = 2
        Public Const MOD_Nada As Integer = 3

        'Constantes para los botones de acción
        'según el ToolBar del formulario Principal frmMain
        Public Const CMD_Salir As Integer = 1
        'Separador
        Public Const CMD_Primero As Integer = 3
        Public Const CMD_Anterior As Integer = 4
        Public Const CMD_Siguiente As Integer = 5
        Public Const CMD_Ultimo As Integer = 6
        'Separador
        Public Const CMD_Nuevo As Integer = 8
        Public Const CMD_Eliminar As Integer = 9
        Public Const CMD_Deshacer As Integer = 10
        Public Const CMD_Grabar As Integer = 11
        'Separador
        Public Const CMD_Pantalla As Integer = 13
        Public Const CMD_Imprimir As Integer = 14
        'Separador
        Public Const CMD_Filtrar As Integer = 16
        Public Const CMD_Buscar As Integer = 17
        Public Const CMD_Divisa As Integer = 18
        'Separador
        Public Const CMD_Ayuda As Integer = 20
        'Separador
        Public Const CMD_Menu As Integer = 22

        Public Const MAX_Botones As Integer = CMD_Menu

        ' --- Constantes de Impresion -------------------
        Public Const CTE_ImpresionPantalla As Integer = 0
        Public Const CTE_ImpresionImpresora As Integer = 1
        Public Const CTE_CancelarImpresion As Integer = 2
        Public compro As Boolean
        Public driv As String
        Public nom As String
        Public port As String
        Public cop As Integer

        '-------------------------------------------------------------------------------------
        'Constantes de mensajes
        Public Const MSG_001 As String = " Grabar los cambios? "
        Public Const MSG_002 As String = " Se regularizarán a 0 las existencias, Continuar? "
        Public Const MSG_003 As String = " Abandonar los cambios? "
        Public Const MSG_004 As String = " Imprimir el Formulario? "
        Public Const MSG_005 As String = " Mensaje Nº 5 "
        Public Const MSG_006 As String = " No existe: "
        Public Const MSG_007 As String = " Se Eliminiaran los Datos Permanentemente. Continuar?"

        Public Const MSG_050 As String = " No se ha podido actualizar el artículo! "
        Public Const MSG_051 As String = " No se ha encontrado el artículo! "
        Public Const MSG_052 As String = " Grabación Realizada! "

        'Errores
        Public Const MSG_100 As String = " Error al grabar los datos! "
        Public Const MSG_101 As String = " Error al borrar los datos! "
        Public Const MSG_102 As String = " Error, el dato está fuera de rango. "
        Public Const MSG_103 As String = " Error en el campo: "
        Public Const MSG_104 As String = " Debe introducir un valor en: "
        Public Const MSG_105 As String = " se desharán los cambios. "

        '-------------------------------------------------------------------------------------
        'Constantes de divisas
        Public Const DIV_Peseta As Integer = 0     'Divisa Peseta
        Public Const DIV_Euro As Integer = 1       'Divisa Euro
        Public Const DEC_Peseta As Integer = 0  'Decimales de trabajo en Pesetas
        Public Const DEC_Euro As Integer = 3    'Decimales de trabajo en Euros

        '*************************************************************************************

        'Constantes para la linea de estado
        Public Const EST_Mensaje As Integer = 1
        Public Const EST_Empresa As Integer = 2
        Public Const EST_Divisa As Integer = 3
        Public Const EST_Usuario As Integer = 4

        Public Const CTE_TiempoEsperaEntornoDatos As Integer = 200
        Public Const CTE_TiempoEsperaTransaccion As Integer = 10
        
        ' Timeout máximo en segundos para pruebas de conexión inicial
        ' Se usa un valor más corto para evitar que la UI se bloquee
        Public Const CTE_TimeoutPruebaConexion As Integer = 10

        Public Const MetrosCubicos As Double = 0.028317 ' Factor de conversión a m3 de 1 pie3


        'Constantes de Tamaño de Etiquetas para Impresion de Bultos (SSCC,Ubicaciones,..)
        Public Const ETI_14Con8 As Integer = 1
        Public Const ETI_12Con9 As Integer = 2

        ' --- Estados de los Grupos ------------------------------
        Public Const EstadoGrupo_Creado As String = "010"
        Public Const EstadoGrupo_Asignado As String = "020"
        Public Const EstadoGrupo_Iniciado As String = "030"
        Public Const EstadoGrupo_Pausado As String = "040"
        Public Const EstadoGrupo_Finalizado As String = "080"
        Public Const EstadoGrupo_Completo As String = "085"
        Public Const EstadoGrupo_Exportado As String = "090"

        ' -- Constantes SSCC
        Public Const IncrementoSerieSSCC_Hipodromo As Integer = 30

    End Module

End Namespace
