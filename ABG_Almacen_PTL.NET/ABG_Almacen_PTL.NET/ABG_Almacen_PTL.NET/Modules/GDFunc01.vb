'*****************************************************************************************
'GDFunc01.vb
'
' Módulo de funciones generales.
' Converted from VB6 to VB.NET
'
'*****************************************************************************************
' FUNCIONES:
' CargaMenu     Función para cargar los menús de la aplicación según la opción elegida
' MenuActivo    Función para activar las opciones correspondientes al menú seleccionado
' CambiaModo    Cambia el modo de la barra de botones del modulo principal frmMain
'=========================================================================================

Imports System
Imports System.Windows.Forms
Imports System.Data
Imports Microsoft.Data.SqlClient
Imports ABG_Almacen_PTL.Classes

Namespace Modules

    Public Module GDFunc01

        ' Recordset Genérico para almacenar los menus y permisos a los que tiene acceso el usuario
        Public r_menu As clGenericaRecordset

        '*****************************************************************************************
        ' CargaMenu: Función para cargar los menús de la aplicación según la opción elegida
        '*****************************************************************************************
        Public Sub CargaMenu(Index As Integer)
            Menu(Index).Nombre = "ABG Almacén RE"
            Select Case Index
                Case CMD_Aduana
                    ReDim Menu(Index).Opcion(2)
                    Menu(Index).Opcion(0).Formulario = ""
                    Menu(Index).Opcion(1).Formulario = ""
                Case CMD_Almacen
                    ReDim Menu(Index).Opcion(2)
                    Menu(Index).Opcion(0).Formulario = "&Reparto Automático"
                    Menu(Index).Opcion(1).Formulario = "&Empaquetado"
            End Select
        End Sub

        '*****************************************************************************************
        ' Función CambiaModo:
        '               Función para activar/desactivar los botones de la barra de herramientas
        '               según el modo de trabajo.
        ' Parámetros:
        '               Modo : Modo de trabajo: MOD_Edición | MOD_Seleccion
        ' Utilización:
        '               Se llama desde los formularios cliente.
        '*****************************************************************************************
        Public Sub CambiaModo(Modo As Integer, tbToolbar As ToolStrip)
            ' Cambia el Modo del toolbar del Formulario Principal MDI: frmMain
            If tbToolbar Is Nothing Then Return

            Select Case Modo
                Case MOD_Seleccion
                    'Entra en el modo de selección de registros
                    SetButtonEnabled(tbToolbar, CMD_Primero, True)
                    SetButtonEnabled(tbToolbar, CMD_Anterior, True)
                    SetButtonEnabled(tbToolbar, CMD_Siguiente, True)
                    SetButtonEnabled(tbToolbar, CMD_Ultimo, True)
                    SetButtonEnabled(tbToolbar, CMD_Nuevo, True)
                    SetButtonEnabled(tbToolbar, CMD_Eliminar, True)
                    SetButtonEnabled(tbToolbar, CMD_Deshacer, False)
                    SetButtonEnabled(tbToolbar, CMD_Grabar, False)
                    SetButtonEnabled(tbToolbar, CMD_Salir, True)
                    SetButtonEnabled(tbToolbar, CMD_Pantalla, True)
                    SetButtonEnabled(tbToolbar, CMD_Imprimir, True)
                    SetButtonEnabled(tbToolbar, CMD_Filtrar, True)
                    SetButtonEnabled(tbToolbar, CMD_Buscar, True)

                Case MOD_Edicion
                    'Entra en el modo de edición de registros
                    SetButtonEnabled(tbToolbar, CMD_Primero, False)
                    SetButtonEnabled(tbToolbar, CMD_Anterior, False)
                    SetButtonEnabled(tbToolbar, CMD_Siguiente, False)
                    SetButtonEnabled(tbToolbar, CMD_Ultimo, False)
                    SetButtonEnabled(tbToolbar, CMD_Nuevo, False)
                    SetButtonEnabled(tbToolbar, CMD_Eliminar, False)
                    SetButtonEnabled(tbToolbar, CMD_Deshacer, True)
                    SetButtonEnabled(tbToolbar, CMD_Grabar, True)
                    SetButtonEnabled(tbToolbar, CMD_Salir, False)
                    SetButtonEnabled(tbToolbar, CMD_Pantalla, False)
                    SetButtonEnabled(tbToolbar, CMD_Imprimir, False)
                    SetButtonEnabled(tbToolbar, CMD_Filtrar, False)
                    SetButtonEnabled(tbToolbar, CMD_Buscar, False)

                Case MOD_Todo
                    'Activa todos los botones
                    For i As Integer = 0 To tbToolbar.Items.Count - 1
                        tbToolbar.Items(i).Enabled = True
                    Next

                Case MOD_Nada
                    'Desactiva todos los botones
                    For i As Integer = 0 To tbToolbar.Items.Count - 1
                        tbToolbar.Items(i).Enabled = False
                    Next
                    If tbToolbar.Items.Count > 0 Then tbToolbar.Items(0).Enabled = True
                    If tbToolbar.Items.Count > MAX_Botones - 1 Then tbToolbar.Items(MAX_Botones - 1).Enabled = True
            End Select
        End Sub

        Private Sub SetButtonEnabled(tbToolbar As ToolStrip, index As Integer, enabled As Boolean)
            If index > 0 AndAlso index <= tbToolbar.Items.Count Then
                tbToolbar.Items(index - 1).Enabled = enabled
            End If
        End Sub

        '**************************************************************************************
        ' Prueba si hay conexión con el Servidor que se le pasa por parámetro
        ' FIEL AL COMPORTAMIENTO VB6: Devuelve True/False sin mostrar mensaje de error
        ' El mensaje de error se muestra en el formulario que llama a esta función
        '**************************************************************************************
        Public Function ProbarConexion(serv As String) As Boolean
            ' En VB6 esta función simplemente intentaba abrir la conexión y
            ' devolvía True si tenía éxito, False si fallaba (sin mostrar mensaje)
            
            ' Si el servidor está vacío, retornar false (igual que en VB6)
            If String.IsNullOrEmpty(serv) Then
                Return False
            End If

            ' Validar que tenemos los parámetros de conexión necesarios
            If String.IsNullOrEmpty(BDDConfig) Then
                Return False
            End If

            ' Construir cadena de conexión similar a VB6:
            ' VB6 usaba: "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=xxx;Password=xxx;Initial Catalog=Config;Data Source=serv;Connect Timeout=30"
            ' En .NET usamos Microsoft.Data.SqlClient con opciones de compatibilidad
            ' Se usa Encrypt=False y TrustServerCertificate=True para compatibilidad con 
            ' servidores SQL Server antiguos (2016 y anteriores) que pueden no tener 
            ' configurado correctamente TLS/SSL
            ' Se usa un timeout de conexión más corto para evitar bloqueos
            ' largos cuando el servidor no está disponible
            Dim timeoutConexion As Integer = Math.Min(BDDTime, CTE_TimeoutPruebaConexion)
            Dim conexion As String = $"Server={serv};Database={BDDConfig};User Id={UsrBDDConfig};Password={UsrKeyConfig};Connect Timeout={timeoutConexion};TrustServerCertificate=True;Encrypt=False;"

            Try
                Cursor.Current = Cursors.WaitCursor
                Application.DoEvents() ' Permitir que la UI se actualice
                
                ' Igual que en VB6: un solo intento de conexión
                Using conn As New SqlConnection(conexion)
                    conn.Open()
                    ' Conexión exitosa
                    Return True
                End Using
                
            Catch ex As Exception
                ' En VB6 simplemente devolvía False sin mostrar mensaje
                ' El error se maneja en el formulario llamante
                Return False
            Finally
                Cursor.Current = Cursors.Default
            End Try
        End Function

        '**************************************************************************************
        'Función:   wfCambiarCadena
        'Objetivo:  Cambiar en la Cadena que se manda, la Cadena Buscada por la Cadena a Reemplazar
        '**************************************************************************************
        Public Function wfCambiarCadena(vCadenaInicial As Object, vCadenaBuscada As Object, vCadenaReemplazar As Object) As Object
            If vCadenaInicial Is Nothing Then Return ""
            Return vCadenaInicial.ToString().Replace(vCadenaBuscada.ToString(), vCadenaReemplazar.ToString())
        End Function

        '**************************************************************************************
        'Función:   wfComaXPunto
        'Objetivo:  Cambiar en la Cadena la Coma por un Punto
        '**************************************************************************************
        Public Function wfComaXPunto(vCadenaInicial As Object) As Object
            Return wfCambiarCadena(vCadenaInicial, ",", ".")
        End Function

        '**************************************************************************************
        'Función:   wfQuitarCorchetes
        'Objetivo:  Elimina los corchetes abierto y cerrado, [ ] , de una cadena
        '**************************************************************************************
        Public Function wfQuitarCorchetes(vCadenaInicial As Object) As Object
            Return wfCambiarCadena(wfCambiarCadena(vCadenaInicial, "[", ""), "]", "")
        End Function

        '**************************************************************************************
        'Función:   wfPonerComillas
        'Objetivo:  Añadir a la Cadena enviada, comillas simples al principio y final
        '**************************************************************************************
        Public Function wfPonerComillas(vCadenaInicial As Object) As Object
            If vCadenaInicial Is Nothing Then Return "''"

            Dim vCadenaTotal As String = vCadenaInicial.ToString().Trim()

            If String.IsNullOrEmpty(vCadenaTotal) Then
                Return "''"
            Else
                ' Primero se sustituyen las posibles comillas simples internas por acentos para evitar error de sintaxis
                vCadenaTotal = vCadenaTotal.Replace("'", "´")
                ' Luego se añaden las comillas simples
                Return $"'{vCadenaTotal}'"
            End If
        End Function

        '**************************************************************************************
        'Función:   wsMensaje
        'Objetivo:  Presentar un Mensaje en Formulario
        '**************************************************************************************
        Public Sub wsMensaje(stMensaje As String, Optional vtTipo As TipoMensaje = TipoMensaje.MENSAJE_Grave)
            Dim icon As MessageBoxIcon
            Select Case vtTipo
                Case TipoMensaje.MENSAJE_Informativo
                    icon = MessageBoxIcon.Information
                Case TipoMensaje.MENSAJE_Exclamacion
                    icon = MessageBoxIcon.Exclamation
                Case Else
                    icon = MessageBoxIcon.Error
            End Select

            MessageBox.Show(stMensaje, "ABG Almacén PTL", MessageBoxButtons.OK, icon)
        End Sub

        '**************************************************************************************
        'Función:   wfColorTerminacion
        'Objetivo:  Obtener los colores para mostrar el grado de terminación de las cantidades
        '**************************************************************************************
        Public Function wfColorTerminacion(vtValor As Object) As System.Drawing.Color
            Dim valor As Double = 0
            If vtValor IsNot Nothing Then
                Double.TryParse(vtValor.ToString(), valor)
            End If

            Select Case valor
                Case 0
                    Return System.Drawing.Color.FromArgb(&HFF, 0, 0)      ' Rojo oscuro
                Case Is >= 100
                    Return System.Drawing.Color.FromArgb(0, &HC0, 0)     ' Verde fuerte
                Case Else
                    Return System.Drawing.Color.FromArgb(&HFF, &H80, 0)  ' Ámbar
            End Select
        End Function

        '**************************************************************************************
        'Función:   ControlEjecucion
        'Objetivo:  Controla una marca de ejecución correcta del programa en el archivo ABG.INI
        '**************************************************************************************
        Public Sub ControlEjecucion()
            Try
                Dim nombreEXE As String = Application.ProductName
                Dim ficINILocal As String = IO.Path.Combine(Application.StartupPath, "ABG.INI")
                GuardarIni(ficINILocal, "Versiones", "Programa", nombreEXE)
                GuardarIni(ficINILocal, "Versiones", "EjecucionCorrecta", "1")
            Catch ex As Exception
                ' Ignore errors
            End Try
        End Sub

        '**************************************************************************************
        'Función:   ActualizaCargador
        'Objetivo:  Actualiza el CargadorABG desde el servidor si hay una versión nueva
        '**************************************************************************************
        Public Sub ActualizaCargador()
            Try
                Dim ficINIlocal As String = IO.Path.Combine(Application.StartupPath, "ABG.INI")
                Dim Ruta As String
                Dim serv As String
                Dim version As String
                Dim version_serv As String

                ' Lectura de la ruta de donde están los programas para actualizar el cargador
                serv = LeerIni(ficINIlocal, "Versiones", "APPServ", "")
                Ruta = LeerIni(ficINIlocal, "Versiones", "RutaProgramas", "")

                If String.IsNullOrEmpty(serv) Then Exit Sub

                If String.IsNullOrEmpty(Ruta) Then
                    Ruta = "\Programas\ABG\Ejecutable\"
                    GuardarIni(ficINIlocal, "Versiones", "RutaProgramas", Ruta)
                End If

                Dim ficINIserv As String = serv & Ruta & "Version.INI"

                ' Se leen las versiones del cargador locales y del servidor
                version_serv = LeerIni(ficINIserv, "CargadorABG", "Version", "")
                version = LeerIni(ficINIlocal, "Versiones", "Cargador", "")

                ' Si las versiones son diferentes hay que actualizar
                If version <> version_serv Then
                    IO.File.Copy(serv & Ruta & "CargadorABG.EXE", IO.Path.Combine(Application.StartupPath, "CargadorABG.EXE"), True)
                    GuardarIni(ficINIlocal, "Versiones", "Cargador", version_serv)
                End If
            Catch ex As Exception
                ' Ignore errors in actualization
            End Try
        End Sub

        '**************************************************************************************
        ' Sobrecarga de CambiaModo sin toolbar (para compatibilidad con VB6)
        '**************************************************************************************
        Public Sub CambiaModo(Modo As Integer)
            ' Versión sin toolbar - no hace nada
            ' En VB6 esto actualizaba el toolbar del frmMain
        End Sub

    End Module

End Namespace
