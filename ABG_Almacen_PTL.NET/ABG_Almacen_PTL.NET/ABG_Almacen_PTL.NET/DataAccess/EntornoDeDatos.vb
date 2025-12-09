'*****************************************************************************************
'EntornoDeDatos.vb
'
' Clase de acceso a datos para la gestión del almacén
' Converted from VB6 DataEnvironment to VB.NET ADO.NET
'*****************************************************************************************

Imports System
Imports System.Data
Imports Microsoft.Data.SqlClient
Imports ABG_Almacen_PTL.Modules

Namespace DataAccess

    Public Class EntornoDeDatos
        Implements IDisposable

        Private _connection As SqlConnection
        Private _disposed As Boolean = False

        ' Cadena de conexión
        Public Property ConnectionString As String

        ' Constructor - usa la cadena de conexión global de Gestión Almacén
        Public Sub New()
            ' Usar la cadena de conexión global si está disponible
            If Not String.IsNullOrEmpty(ConexionGestionAlmacen) Then
                ConnectionString = ConexionGestionAlmacen
            Else
                ConnectionString = ""
            End If
        End Sub

        ' Constructor con cadena de conexión
        Public Sub New(connectionString As String)
            Me.ConnectionString = connectionString
        End Sub

        ' Abrir conexión
        Public Sub Open(Optional connectionString As String = Nothing)
            If Not String.IsNullOrEmpty(connectionString) Then
                Me.ConnectionString = connectionString
            End If

            If _connection Is Nothing Then
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

        ' Ejecutar consulta y devolver DataTable
        Public Function Execute(sql As String) As DataTable
            If _connection Is Nothing OrElse _connection.State <> ConnectionState.Open Then
                Open()
            End If

            Dim dt As New DataTable()
            Using cmd As New SqlCommand(sql, _connection)
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos
                Using adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
            Return dt
        End Function

        ' Ejecutar comando sin resultado
        Public Function ExecuteNonQuery(sql As String) As Integer
            If _connection Is Nothing OrElse _connection.State <> ConnectionState.Open Then
                Open()
            End If

            Using cmd As New SqlCommand(sql, _connection)
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos
                Return cmd.ExecuteNonQuery()
            End Using
        End Function

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

        ' Asignación de BAC a Tablilla
        Public Sub AsignacionBacATablilla(bac As String, grupo As Integer, tablilla As Integer, usuario As Integer)
            Using cmd As New SqlCommand("dbo.AsignacionBacATablilla", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@bac", bac)
                cmd.Parameters.AddWithValue("@grupo", grupo)
                cmd.Parameters.AddWithValue("@tablilla", tablilla)
                cmd.Parameters.AddWithValue("@usuario", usuario)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ' Insertar detalle de BAC
        Public Sub InsertaDetalleBac(bac As String, grupo As Integer, tablilla As Integer,
                                     articulo As Integer, cantidad As Integer, usuario As Integer)
            Using cmd As New SqlCommand("dbo.InsertaDetalleBac", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@bac", bac)
                cmd.Parameters.AddWithValue("@grupo", grupo)
                cmd.Parameters.AddWithValue("@tablilla", tablilla)
                cmd.Parameters.AddWithValue("@articulo", articulo)
                cmd.Parameters.AddWithValue("@cantidad", cantidad)
                cmd.Parameters.AddWithValue("@usuario", usuario)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ' Insertar log de empaquetado
        Public Sub InsertaLogEmpaquetado(grupo As Integer, tablilla As Integer, caja As Integer,
                                          codigo As String, articulo As Integer, cantidad As Integer,
                                          tipo As Integer, bac As String, idMov As Integer,
                                          descripcion As String, observacion As String,
                                          puesto As Integer, usuario As Integer)
            Using cmd As New SqlCommand("dbo.InsertaLogEmpaquetado", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@Grupo", grupo)
                cmd.Parameters.AddWithValue("@Tablilla", tablilla)
                cmd.Parameters.AddWithValue("@Caja", caja)
                cmd.Parameters.AddWithValue("@Codigo", codigo)
                cmd.Parameters.AddWithValue("@Articulo", articulo)
                cmd.Parameters.AddWithValue("@Cantidad", cantidad)
                cmd.Parameters.AddWithValue("@Tipo", tipo)
                cmd.Parameters.AddWithValue("@BAC", bac)
                cmd.Parameters.AddWithValue("@IdMov", idMov)
                cmd.Parameters.AddWithValue("@Descripcion", descripcion)
                cmd.Parameters.AddWithValue("@Observacion", observacion)
                cmd.Parameters.AddWithValue("@Puesto", puesto)
                cmd.Parameters.AddWithValue("@Usuario", usuario)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ' Dame grupos
        Public Function DameGrupos() As DataTable
            Return ExecuteStoredProcedure("dbo.DameGrupos")
        End Function

        ' Dame fecha y hora del sistema
        Public Function DameFechaHoraHoy() As Date
            Dim dt As DataTable = ExecuteStoredProcedure("dbo.DameFechaHoraHoy")
            If dt.Rows.Count > 0 Then
                Return CDate(dt.Rows(0)("Hoy"))
            End If
            Return DateTime.Now
        End Function

        ' Dame tipos de cajas activas
        Public Function DameTiposCajasActivas() As DataTable
            Return ExecuteStoredProcedure("dbo.DameTiposCajasActivas")
        End Function

        ' Dame tablillas de grupo
        Public Function DameTablillasDeGrupo(grupo As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameTablillasGrupo",
                New SqlParameter("@Grupo", grupo))
        End Function

        ' Dame cajas de grupo y tablilla
        Public Function DameCajasDeGrupoTablilla(grupo As Integer, tablilla As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameCajasPorGrupoTablilla",
                New SqlParameter("@Grupo", grupo),
                New SqlParameter("@Tablilla", tablilla))
        End Function

        ' Crear caja grupo tablilla PTL
        Public Sub CrearCajaGrupoTablillaPTL(grupo As Integer, tablilla As Integer, caja As String,
                                              tipo As Integer, sscc As String, bac As String)
            Using cmd As New SqlCommand("dbo.CrearCajaGrupoTablillaPTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@inGrupo", grupo)
                cmd.Parameters.AddWithValue("@inTablilla", tablilla)
                cmd.Parameters.AddWithValue("@stCaja", caja)
                cmd.Parameters.AddWithValue("@inTipo", tipo)
                cmd.Parameters.AddWithValue("@SSCC", sscc)
                cmd.Parameters.AddWithValue("@BAC", bac)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ' Dame datos de caja PTL
        Public Function DameDatosCAJAdePTL(sscc As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameDatosCAJAdePTL",
                New SqlParameter("@SSCC", sscc))
        End Function

        ' Traspasa BAC a CAJA de PTL
        Public Function TraspasaBACaCAJAdePTL(bac As String, usuario As Integer) As (SSCC As String, Retorno As Integer, Mensaje As String)
            Using cmd As New SqlCommand("dbo.TraspasaBACaCAJAdePTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@BAC", bac)
                cmd.Parameters.AddWithValue("@Usuario", usuario)

                Dim paramSSCC As New SqlParameter("@SSCC", SqlDbType.VarChar, 50)
                paramSSCC.Direction = ParameterDirection.InputOutput
                paramSSCC.Value = ""
                cmd.Parameters.Add(paramSSCC)

                Dim paramRetorno As New SqlParameter("@Retorno", SqlDbType.SmallInt)
                paramRetorno.Direction = ParameterDirection.InputOutput
                paramRetorno.Value = 0
                cmd.Parameters.Add(paramRetorno)

                Dim paramMsg As New SqlParameter("@msgSalida", SqlDbType.VarChar, 1024)
                paramMsg.Direction = ParameterDirection.InputOutput
                paramMsg.Value = ""
                cmd.Parameters.Add(paramMsg)

                cmd.ExecuteNonQuery()

                Return (paramSSCC.Value.ToString(),
                        CInt(paramRetorno.Value),
                        paramMsg.Value.ToString())
            End Using
        End Function

        ' Dame contenido de caja
        Public Function DameContenidoCajaGrupo(grupo As Integer, tablilla As Integer, caja As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameContenidoCajaGrupo",
                New SqlParameter("@Grupo", grupo),
                New SqlParameter("@Tablilla", tablilla),
                New SqlParameter("@CAJA", caja))
        End Function

        ' Dame puestos de trabajo PTL
        Public Function DamePuestosTrabajoPTL() As DataTable
            Return ExecuteStoredProcedure("dbo.DamePuestosTrabajoPTL")
        End Function

        ' Dame datos BAC de PTL
        Public Function DameDatosBACdePTL(bac As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameDatosBACdePTL",
                New SqlParameter("@BAC", bac))
        End Function

        ' Dame contenido BAC de Grupo
        Public Function DameContenidoBacGrupo(grupo As Integer, bac As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameContenidoBacGrupo",
                New SqlParameter("@Grupo", grupo),
                New SqlParameter("@BAC", bac))
        End Function

        ' Dame última caja de BAC
        Public Function DameUltimaCajaDeBAC(bac As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameUltimaCajaDeBAC",
                New SqlParameter("@BAC", bac))
        End Function

        ' Dame caja Grupo Tablilla PTL
        Public Function DameCajaGrupoTablillaPTL(grupo As Integer, tablilla As Integer, caja As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameCajaGrupoTablillaPTL",
                New SqlParameter("@Grupo", grupo),
                New SqlParameter("@Tablilla", tablilla),
                New SqlParameter("@Caja", caja))
        End Function

        ' Dame cajas Grupo Tablilla PTL
        Public Function DameCajasGrupoTablillaPTL(grupo As Integer, tablilla As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameCajasGrupoTablillaPTL",
                New SqlParameter("@Grupo", grupo),
                New SqlParameter("@Tablilla", tablilla))
        End Function

        ' Cambiar estado BAC de PTL
        Public Sub CambiaEstadoBACdePTL(bac As String, estado As Integer, usuario As Integer, ByRef retorno As Integer, ByRef msgSalida As String)
            EnsureConnectionOpen()
            Using cmd As New SqlCommand("dbo.CambiaEstadoBACdePTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos
                cmd.Parameters.AddWithValue("@BAC", bac)
                cmd.Parameters.AddWithValue("@Estado", estado)
                cmd.Parameters.AddWithValue("@Usuario", usuario)

                Dim paramRetorno As New SqlParameter("@Retorno", SqlDbType.SmallInt)
                paramRetorno.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramRetorno)

                Dim paramMsg As New SqlParameter("@msgSalida", SqlDbType.VarChar, 1024)
                paramMsg.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramMsg)

                cmd.ExecuteNonQuery()

                retorno = If(paramRetorno.Value IsNot DBNull.Value, CInt(paramRetorno.Value), -1)
                msgSalida = If(paramMsg.Value IsNot DBNull.Value, paramMsg.Value.ToString(), "")
            End Using
        End Sub

        ' Retirar BAC de PTL
        Public Sub RetirarBACdePTL(bac As String, usuario As Integer, ByRef retorno As Integer, ByRef msgSalida As String)
            EnsureConnectionOpen()
            Using cmd As New SqlCommand("dbo.RetirarBACdePTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos
                cmd.Parameters.AddWithValue("@BAC", bac)
                cmd.Parameters.AddWithValue("@Usuario", usuario)

                Dim paramRetorno As New SqlParameter("@Retorno", SqlDbType.SmallInt)
                paramRetorno.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramRetorno)

                Dim paramMsg As New SqlParameter("@msgSalida", SqlDbType.VarChar, 1024)
                paramMsg.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramMsg)

                cmd.ExecuteNonQuery()

                retorno = If(paramRetorno.Value IsNot DBNull.Value, CInt(paramRetorno.Value), -1)
                msgSalida = If(paramMsg.Value IsNot DBNull.Value, paramMsg.Value.ToString(), "")
            End Using
        End Sub

        ' Actualiza caja BAC PTL
        Public Sub ActualizaCajaBACPTL(bac As String, caja As String)
            EnsureConnectionOpen()
            Using cmd As New SqlCommand("dbo.ActualizaCajaBACPTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos
                cmd.Parameters.AddWithValue("@BAC", bac)
                cmd.Parameters.AddWithValue("@Caja", caja)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ' Cambiar tipo de caja PTL
        Public Sub CambiaTipoCajaPTL(tipoCaja As Integer, bac As String, sscc As String, usuario As Integer)
            EnsureConnectionOpen()
            Using cmd As New SqlCommand("dbo.CambiaTipoCajaPTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos
                cmd.Parameters.AddWithValue("@TipoCaja", tipoCaja)
                cmd.Parameters.AddWithValue("@BAC", If(String.IsNullOrEmpty(bac), DBNull.Value, bac))
                cmd.Parameters.AddWithValue("@SSCC", If(String.IsNullOrEmpty(sscc), DBNull.Value, sscc))
                cmd.Parameters.AddWithValue("@Usuario", usuario)
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        ' Combinar cajas PTL
        Public Sub CombinarCajasPTL(sscc1 As String, sscc2 As String, usuario As Integer, ByRef retorno As Integer, ByRef msgSalida As String)
            EnsureConnectionOpen()
            Using cmd As New SqlCommand("dbo.CombinarCajasPTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos
                cmd.Parameters.AddWithValue("@SSCC1", sscc1)
                cmd.Parameters.AddWithValue("@SSCC2", sscc2)
                cmd.Parameters.AddWithValue("@Usuario", usuario)

                Dim paramRetorno As New SqlParameter("@Retorno", SqlDbType.SmallInt)
                paramRetorno.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramRetorno)

                Dim paramMsg As New SqlParameter("@msgSalida", SqlDbType.VarChar, 1024)
                paramMsg.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramMsg)

                cmd.ExecuteNonQuery()

                retorno = If(paramRetorno.Value IsNot DBNull.Value, CInt(paramRetorno.Value), -1)
                msgSalida = If(paramMsg.Value IsNot DBNull.Value, paramMsg.Value.ToString(), "")
            End Using
        End Sub

        ' Traspasa BAC a CAJA de PTL (versión con ByRef - igual que VB6)
        Public Sub TraspasaBACaCAJAdePTLByRef(bac As String, usuario As Integer, ssccBase As String, ByRef retorno As Integer, ByRef msgSalida As String)
            EnsureConnectionOpen()
            Using cmd As New SqlCommand("dbo.TraspasaBACaCAJAdePTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos
                cmd.Parameters.AddWithValue("@BAC", bac)
                cmd.Parameters.AddWithValue("@Usuario", usuario)

                Dim paramSSCC As New SqlParameter("@SSCC", SqlDbType.VarChar, 50)
                paramSSCC.Direction = ParameterDirection.InputOutput
                paramSSCC.Value = ssccBase
                cmd.Parameters.Add(paramSSCC)

                Dim paramRetorno As New SqlParameter("@Retorno", SqlDbType.SmallInt)
                paramRetorno.Direction = ParameterDirection.InputOutput
                paramRetorno.Value = 0
                cmd.Parameters.Add(paramRetorno)

                Dim paramMsg As New SqlParameter("@msgSalida", SqlDbType.VarChar, 1024)
                paramMsg.Direction = ParameterDirection.InputOutput
                paramMsg.Value = ""
                cmd.Parameters.Add(paramMsg)

                cmd.ExecuteNonQuery()

                retorno = If(paramRetorno.Value IsNot DBNull.Value, CInt(paramRetorno.Value), -1)
                msgSalida = If(paramMsg.Value IsNot DBNull.Value, paramMsg.Value.ToString(), "")
            End Using
        End Sub

        ' Cambiar unidades de artículo en caja PTL
        Public Sub CambiaUnidadesArtCajaPTL(sscc As String, articulo As Integer, cantidad As Integer, usuario As Integer, ByRef retorno As Integer, ByRef msgSalida As String)
            EnsureConnectionOpen()
            Using cmd As New SqlCommand("dbo.CambiaUnidadesArtCajaPTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = Modules.CTE_TiempoEsperaEntornoDatos
                cmd.Parameters.AddWithValue("@SSCC", sscc)
                cmd.Parameters.AddWithValue("@Articulo", articulo)
                cmd.Parameters.AddWithValue("@Cantidad", cantidad)
                cmd.Parameters.AddWithValue("@Usuario", usuario)

                Dim paramRetorno As New SqlParameter("@Retorno", SqlDbType.SmallInt)
                paramRetorno.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramRetorno)

                Dim paramMsg As New SqlParameter("@msgSalida", SqlDbType.VarChar, 1024)
                paramMsg.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramMsg)

                cmd.ExecuteNonQuery()

                retorno = If(paramRetorno.Value IsNot DBNull.Value, CInt(paramRetorno.Value), -1)
                msgSalida = If(paramMsg.Value IsNot DBNull.Value, paramMsg.Value.ToString(), "")
            End Using
        End Sub

        ' Método auxiliar para asegurar que la conexión esté abierta
        Private Sub EnsureConnectionOpen()
            If _connection Is Nothing OrElse _connection.State <> ConnectionState.Open Then
                Open()
            End If
        End Sub

        ' Dame artículo por código (para consulta)
        Public Function DameArticuloConsulta(articulo As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameArticuloConsulta",
                New SqlParameter("@Articulo", articulo))
        End Function

        ' Dame artículo por EAN13
        Public Function DameArticuloEAN13(ean13 As String) As DataTable
            Return ExecuteStoredProcedure("dbo.DameArticuloEAN13",
                New SqlParameter("@EAN13", ean13))
        End Function

        ' Reserva BAC de PTL para artículo
        Public Sub ReservaBACdePTL(articulo As Long, usuario As Integer, ByRef retorno As Integer, ByRef msgSalida As String)
            EnsureConnectionOpen()
            Using cmd As New SqlCommand("dbo.ReservaBACdePTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = CTE_TiempoEsperaEntornoDatos
                cmd.Parameters.AddWithValue("@Articulo", articulo)
                cmd.Parameters.AddWithValue("@Usuario", usuario)

                Dim paramRetorno As New SqlParameter("@Retorno", SqlDbType.SmallInt)
                paramRetorno.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramRetorno)

                Dim paramMsg As New SqlParameter("@msgSalida", SqlDbType.VarChar, 1024)
                paramMsg.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramMsg)

                cmd.ExecuteNonQuery()

                retorno = If(paramRetorno.Value IsNot DBNull.Value, CInt(paramRetorno.Value), -1)
                msgSalida = If(paramMsg.Value IsNot DBNull.Value, paramMsg.Value.ToString(), "")
            End Using
        End Sub

        ' Dame datos de ubicación PTL
        Public Function DameDatosUbicacionPTL(alf As Integer, alm As Integer, blo As Integer, fil As Integer, alt As Integer) As DataTable
            Return ExecuteStoredProcedure("dbo.DameDatosUbicacionPTL",
                New SqlParameter("@ALF", alf),
                New SqlParameter("@ALM", alm),
                New SqlParameter("@BLO", blo),
                New SqlParameter("@FIL", fil),
                New SqlParameter("@ALT", alt))
        End Function

        ' Consulta BAC de PTL (existencia en GAUBIBAC)
        Public Function ConsultaBACdePTL(bac As String) As DataTable
            Return ExecuteStoredProcedure("dbo.ConsultaBACdePTL",
                New SqlParameter("@BAC", bac))
        End Function

        ' Ubicar BAC en PTL
        Public Sub UbicarBACenPTL(bac As String, ubicacion As Integer, usuario As Integer, ByRef retorno As Integer, ByRef msgSalida As String)
            EnsureConnectionOpen()
            Using cmd As New SqlCommand("dbo.UbicarBACenPTL", _connection)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = CTE_TiempoEsperaEntornoDatos
                cmd.Parameters.AddWithValue("@BAC", bac)
                cmd.Parameters.AddWithValue("@Ubicacion", ubicacion)
                cmd.Parameters.AddWithValue("@Usuario", usuario)

                Dim paramRetorno As New SqlParameter("@Retorno", SqlDbType.SmallInt)
                paramRetorno.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramRetorno)

                Dim paramMsg As New SqlParameter("@msgSalida", SqlDbType.VarChar, 1024)
                paramMsg.Direction = ParameterDirection.Output
                cmd.Parameters.Add(paramMsg)

                cmd.ExecuteNonQuery()

                retorno = If(paramRetorno.Value IsNot DBNull.Value, CInt(paramRetorno.Value), -1)
                msgSalida = If(paramMsg.Value IsNot DBNull.Value, paramMsg.Value.ToString(), "")
            End Using
        End Sub

        ' Propiedad para acceder a la conexión directamente (compatibilidad)
        Public ReadOnly Property GestionAlmacen As SqlConnection
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

        ' Numerador SSCC Hipodromo
        Public Function DameNumeradorSSCCHipodromo() As DataTable
            Return ExecuteStoredProcedure("dbo.DameNumeradorSSCCHipodromo")
        End Function

        Public Sub ActualizaNumeradorSSCCHipodromo(numerador As Integer)
            ExecuteStoredProcedure("dbo.ActualizaNumeradorSSCCHipodromo",
                                   New SqlParameter("@Numerador", numerador))
        End Sub

        Public Sub InsertaHistoricoSSCCHipodromo(tipo As Integer, sscc As String, descripcion As String, grupo As Integer, tablilla As Integer)
            ExecuteStoredProcedure("dbo.InsertaHistoricoSSCCHipodromo",
                                   New SqlParameter("@Tipo", tipo),
                                   New SqlParameter("@SSCC", sscc),
                                   New SqlParameter("@Descripcion", descripcion),
                                   New SqlParameter("@Grupo", grupo),
                                   New SqlParameter("@Tablilla", tablilla))
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
