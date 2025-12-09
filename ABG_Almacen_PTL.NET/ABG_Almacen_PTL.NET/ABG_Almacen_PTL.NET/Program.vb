'*****************************************************************************************
' ABG Almacén PTL - Programa Principal
' Converted from VB6 to VB.NET
'*****************************************************************************************

Imports System.Data
Imports System.Windows.Forms
Imports ABG_Almacen_PTL.Modules
Imports ABG_Almacen_PTL.DataAccess

Friend Module Program

    <STAThread()>
    Friend Sub Main(args As String())
        Application.SetHighDpiMode(HighDpiMode.SystemAware)
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        ' Inicializar variables globales
        InicializarAplicacion()

        ' Mostrar formulario de login
        Dim frmLogin As New frmInicio()
        If frmLogin.ShowDialog() = DialogResult.OK AndAlso frmLogin.LoginSucceeded Then
            ' Cargar configuración de empresa después del login exitoso
            CargarConfiguracionEmpresa()

            ' Ejecutar la aplicación principal
            Application.Run(New frmMain())
        Else
            ' Login cancelado o fallido
            Application.Exit()
        End If
    End Sub

    ''' <summary>
    ''' Carga la configuración de la empresa después del login exitoso
    ''' Similar a CargarParametrosEmpresa en VB6
    ''' </summary>
    Private Sub CargarConfiguracionEmpresa()
        Try
            Using edC As New edConfig()
                edC.Open()

                ' Obtener parámetros de empresa
                Dim dtParams As DataTable = edC.DameParametrosEmpresa(CodEmpresa)

                If dtParams.Rows.Count > 0 Then
                    Dim row As DataRow = dtParams.Rows(0)

                    ' Configurar EmpresaTrabajo
                    If EmpresaTrabajo Is Nothing Then
                        EmpresaTrabajo = New DatosEmpresaTrabajo()
                    End If

                    EmpresaTrabajo.Codigo = CodEmpresa

                    ' Nombre de empresa
                    If Not IsDBNull(row("empnom")) Then
                        EmpresaTrabajo.Nombre = row("empnom").ToString()
                        Empresa = EmpresaTrabajo.Nombre
                    End If

                    ' Servidor de Gestión de Almacén
                    If Not IsDBNull(row("empsga")) Then
                        EmpresaTrabajo.Servidor = row("empsga").ToString().Replace("[", "").Replace("]", "")
                    End If

                    ' Base de Datos de Gestión de Almacén
                    If Not IsDBNull(row("empbga")) Then
                        EmpresaTrabajo.BaseDeDatos = row("empbga").ToString()
                    End If

                    ' Usuario de Gestión de Almacén
                    If Not IsDBNull(row("empuga")) Then
                        EmpresaTrabajo.Usuario = row("empuga").ToString()
                    End If

                    ' Clave de Gestión de Almacén
                    If Not IsDBNull(row("empkga")) Then
                        EmpresaTrabajo.Contrasena = row("empkga").ToString()
                    End If

                    ' EAN de empresa
                    If Not IsDBNull(row("empean")) Then
                        EmpresaTrabajo.CodigoEAN = row("empean").ToString()
                    End If

                    ' Ruta de informes
                    If Not IsDBNull(row("emprin")) Then
                        EmpresaTrabajo.RutaInformes = row("emprin").ToString()
                        wRInformes = EmpresaTrabajo.RutaInformes
                    End If

                    ' Construir cadena de conexión para Gestión de Almacén
                    ' Similar a ConexionGestionAlmacen en VB6
                    If Not String.IsNullOrEmpty(EmpresaTrabajo.Servidor) AndAlso
                       Not String.IsNullOrEmpty(EmpresaTrabajo.BaseDeDatos) Then

                        ConexionGestionAlmacen = $"Server={EmpresaTrabajo.Servidor};" &
                                                 $"Database={EmpresaTrabajo.BaseDeDatos};" &
                                                 $"User Id={EmpresaTrabajo.Usuario};" &
                                                 $"Password={EmpresaTrabajo.Contrasena};" &
                                                 $"Connect Timeout={BDDTime};" &
                                                 $"TrustServerCertificate=True;Encrypt=False;"
                    Else
                        ' Usar la conexión de configuración por defecto si no hay datos de Gestión Almacén
                        ConexionGestionAlmacen = ConexionConfig
                    End If
                End If

                edC.Close()
            End Using

            ' Obtener nombre del PC
            nombrePC = Environment.MachineName

        Catch ex As Exception
            MessageBox.Show($"Error al cargar configuración de empresa: {ex.Message}",
                          "Error de Configuración",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Warning)
            ' Usar conexión de configuración como fallback
            ConexionGestionAlmacen = ConexionConfig
        End Try
    End Sub

    ''' <summary>
    ''' Inicializa las variables globales de la aplicación
    ''' </summary>
    Private Sub InicializarAplicacion()
        Try
            ' Leer configuración del archivo INI
            ficINI = IO.Path.Combine(Application.StartupPath, "abg.ini")

            ' Si no existe el archivo INI, crearlo
            If Not IO.File.Exists(ficINI) Then
                CrearABGIni(ficINI)
            End If

            ' Leer parámetros del INI (igual que en VB6)
            BDDServ = LeerIni(ficINI, "Conexion", "BDDServ", "")
            BDDServLocal = LeerIni(ficINI, "Conexion", "BDDServLocal", "")
            
            ' Validar timeout - usar valor por defecto si no es válido
            Dim timeoutStr As String = LeerIni(ficINI, "Conexion", "BDDTime", "30")
            If Not Integer.TryParse(timeoutStr, BDDTime) OrElse BDDTime < 5 Then
                BDDTime = 30  ' Valor por defecto
            End If
            
            BDDConfig = LeerIni(ficINI, "Conexion", "BDDConfig", "Config")
            
            ' Si BDDConfig está vacío, usar valor por defecto
            If String.IsNullOrEmpty(BDDConfig) Then
                BDDConfig = "Config"
            End If
            
            UsrBDDConfig = "ABG"  ' El usuario es fijo
            UsrKeyConfig = "A_34ggyx4"    ' Su contraseña también

            ' Leer varios
            UsrDefault = LeerIni(ficINI, "Varios", "UsrDefault", "")
            
            Dim empDefaultStr As String = LeerIni(ficINI, "Varios", "EmpDefault", "0")
            If Not Integer.TryParse(empDefaultStr, CodEmpresa) Then
                CodEmpresa = 0
            End If
            
            Dim pueDefaultStr As String = LeerIni(ficINI, "Varios", "PueDefault", "1")
            If Not Integer.TryParse(pueDefaultStr, wPuestoTrabajo.Id) Then
                wPuestoTrabajo.Id = 1
            End If

            ' Construir cadena de conexión para la base de datos de configuración
            ' Se usa Encrypt=False y TrustServerCertificate=True para compatibilidad con 
            ' servidores SQL Server antiguos (2016 y anteriores) que pueden no tener 
            ' configurado correctamente TLS/SSL
            If Not String.IsNullOrEmpty(BDDServLocal) Then
                ConexionConfig = $"Server={BDDServLocal};Database={BDDConfig};User Id={UsrBDDConfig};Password={UsrKeyConfig};Connect Timeout={BDDTime};TrustServerCertificate=True;Encrypt=False;"
            Else
                ConexionConfig = ""
            End If

            ' Control de ejecución
            ControlEjecucion()

        Catch ex As Exception
            MessageBox.Show($"Error al inicializar la aplicación: {ex.Message}" & vbCrLf & vbCrLf &
                          "Por favor, verifique que el archivo abg.ini existe y tiene una configuración válida.",
                          "Error de Inicialización",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Crea el fichero de inicio ABG.INI
    ''' </summary>
    Private Sub CrearABGIni(fichero As String)
        ' Configuración de la pantalla
        GuardarIni(fichero, "Pantalla", "MainLeft", "-60")
        GuardarIni(fichero, "Pantalla", "MainTop", "-60")
        GuardarIni(fichero, "Pantalla", "MainWidth", "15480")
        GuardarIni(fichero, "Pantalla", "MainHeight", "11220")

        ' Conexión
        GuardarIni(fichero, "Conexion", "BDDTime", "30")
        GuardarIni(fichero, "Conexion", "BDDConfig", "Config")

        ' Servidores por defecto si no existe el ABG.INI
        GuardarIni(fichero, "Conexion", "BDDServ", "SELENE")
        GuardarIni(fichero, "Conexion", "BDDServLocal", "GROOT")

        ' Varios
        GuardarIni(fichero, "Varios", "wDirExport", "")
        GuardarIni(fichero, "Varios", "UsrDefault", "")
        GuardarIni(fichero, "Varios", "EmpDefault", "")
        GuardarIni(fichero, "Varios", "PueDefault", "")
    End Sub

End Module
