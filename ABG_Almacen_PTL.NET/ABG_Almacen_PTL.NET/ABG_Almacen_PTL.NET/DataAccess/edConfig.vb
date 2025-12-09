'*****************************************************************************************
'edConfig.vb
'
' Clase de acceso a datos para la configuración del sistema
' Converted from VB6 DataEnvironment to VB.NET ADO.NET
'*****************************************************************************************

Imports System
Imports System.Data
Imports Microsoft.Data.SqlClient

Namespace DataAccess

    Public Class edConfig
        Implements IDisposable

        Private _connection As SqlConnection
        Private _disposed As Boolean = False

        ' Cadena de conexión
        Public Property ConnectionString As String

        ' Constructor
        Public Sub New()
            ConnectionString = ""
        End Sub

        ' Constructor con cadena de conexión
        Public Sub New(connectionString As String)
            Me.ConnectionString = connectionString
        End Sub

        ' Abrir conexión
        Public Sub Open(Optional connectionString As String = Nothing)
            If Not String.IsNullOrEmpty(connectionString) Then
                Me.ConnectionString = connectionString
            ElseIf String.IsNullOrEmpty(Me.ConnectionString) Then
                ' Usar la cadena de conexión de configuración global
                Me.ConnectionString = Modules.ConexionConfig
            End If

            If _connection Is Nothing Then
                _connection = New SqlConnection(Me.ConnectionString)
            ElseIf _connection.ConnectionString <> Me.ConnectionString Then
                ' Si la cadena de conexión cambió, cerrar y recrear
                If _connection.State = ConnectionState.Open Then
                    _connection.Close()
                End If
                _connection.Dispose()
                _connection = New SqlConnection(Me.ConnectionString)
            End If

            If _connection.State <> ConnectionState.Open Then
                _connection.Open()
            End If
        End Sub

        ' Cerrar conexión
        Public Sub Close()
            If _connection IsNot Nothing AndAlso _connection.State = ConnectionState.Open Then
                _connection.Close()
            End If
        End Sub

        ' Ejecutar stored procedure
        Public Function ExecuteStoredProcedure(procedureName As String, ParamArray parameters() As SqlParameter) As DataTable
            If _connection Is Nothing OrElse _connection.State <> ConnectionState.Open Then
                Open()
            End If

            Dim dt As New DataTable()
            Using cmd As New SqlCommand(procedureName, _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos

                If parameters IsNot Nothing Then
                    For Each param As SqlParameter In parameters
                        cmd.Parameters.Add(param)
                    Next
                End If

                Using adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
            Return dt
        End Function

        ' --- Métodos específicos del sistema ---

        ' Busca usuario por nombre
        Public Function BuscaUsuario(nombre As String) As DataTable
            Return ExecuteStoredProcedure("dbo.BuscaUsuario",
                New SqlParameter("@Nombre", nombre))
        End Function

        ' Dame menús del usuario
        Public Function DameMenusUsuario(usuario As String, empresa As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameMenusUsuario",
                New SqlParameter("@Usuario", usuario),
                New SqlParameter("@emp", empresa))
        End Function

        ' Dame usuario por ID
        Public Function DameUsuarioPorId(id As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameUsuarioPorId",
                New SqlParameter("@Id", id))
        End Function

        ' Dame parámetros de empresa
        Public Function DameParametrosEmpresa(empresa As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameParametrosEmpresa",
                New SqlParameter("@emp", empresa))
        End Function

        ' Dame código de empresa
        Public Function DameCodigoEmpresa(nombre As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameCodigoEmpresa",
                New SqlParameter("@nom", nombre))
        End Function

        ' Dame empresas por grupo
        Public Function DameEmpresasPorGrupo(grupo As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameEmpresasPorGrupo",
                New SqlParameter("@grp", grupo))
        End Function

        ' Dame empresas acceso usuario
        Public Function DameEmpresasAccesoUsuario(usuario As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameEmpresasAccesoUsuario",
                New SqlParameter("@usr", usuario))
        End Function

        ' Dame grupos por usuario
        Public Function DameGruposPorUsuario(userId As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameGruposPorUsuario",
                New SqlParameter("@uid", userId))
        End Function

        ' Dame puesto de trabajo
        Public Function DamePuestoTrabajo(codigo As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DamePuestoTrabajo",
                New SqlParameter("@Codigo", codigo))
        End Function

        ' Dame código de puesto
        Public Function DameCodigoPuesto(descripcion As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameCodigoPuesto",
                New SqlParameter("@DescripcionPuesto", descripcion))
        End Function

        ' Dame puestos
        Public Function DamePuestos() As DataTable
            Return ExecuteStoredProcedure("dbo.DamePuestos")
        End Function

        ' Dame informes
        Public Function DameInformes() As DataTable
            Return ExecuteStoredProcedure("dbo.DameInformes")
        End Function

        ' Dame empresa FTP
        Public Function DameEmpresaFTP(empresa As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameEmpresaFTP",
                New SqlParameter("@Empresa", empresa))
        End Function

        ' Tiene permiso sobre informe
        Public Function TienePermisoSobreInforme(usuario As Integer, idInforme As Integer) As Integer
            Using cmd As New SqlCommand("dbo.TienePermisoSobreInforme", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@Usuario", usuario)
                cmd.Parameters.AddWithValue("@Id_Informe", idInforme)

                Dim paramCuantos As New SqlParameter("@Cuantos", SqlDbType.Int)
                paramCuantos.Direction = ParameterDirection.InputOutput
                paramCuantos.Value = 0
                cmd.Parameters.Add(paramCuantos)

                cmd.ExecuteNonQuery()

                Return CInt(paramCuantos.Value)
            End Using
        End Function

        ' Propiedad para acceder a la conexión directamente (compatibilidad)
        Public ReadOnly Property Config As SqlConnection
            Get
                If _connection Is Nothing Then
                    _connection = New SqlConnection(ConnectionString)
                End If
                Return _connection
            End Get
        End Property

        ' Implementación de IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    If _connection IsNot Nothing Then
                        If _connection.State = ConnectionState.Open Then
                            _connection.Close()
                        End If
                        _connection.Dispose()
                        _connection = Nothing
                    End If
                End If
                _disposed = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Dispose(False)
        End Sub

    End Class

End Namespace
