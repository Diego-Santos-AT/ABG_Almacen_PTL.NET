'*****************************************************************************************
'GDFunc02.vb
'
' Módulo de funciones de relación de datos
' Converted from VB6 to VB.NET
'
'*****************************************************************************************

Imports System
Imports System.Data
Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms

Namespace Modules

    Public Module GDFunc02

        '**************************************************************************************
        'Función:   wfBuscarEnArray
        'Objetivo:  Buscar en un array de dos dimensiones por la primera para devolver el valor de la segunda
        '**************************************************************************************
        Public Function wfBuscarEnArray(vArray As Object()(), vValorBuscar As Object) As Object
            'Si no encuentra nada devuelve ""
            If vArray Is Nothing Then Return ""

            For Each item As Object() In vArray
                If item IsNot Nothing AndAlso item.Length >= 2 Then
                    If item(0) IsNot Nothing AndAlso item(0).ToString() = vValorBuscar.ToString() Then
                        Return item(1)
                    End If
                End If
            Next

            Return ""
        End Function

        '**************************************************************************************
        ' Procedimientos de Impresión de Etiquetas de Cajas
        '**************************************************************************************
        Public Sub wsImprimirEtiquetasCajas(vtDatosEtiquetasCajas As Object(,), blPantalla As Boolean, MOD_Nombre As String)
            ' En VB.NET moderno, la impresión se manejaría diferente
            ' Este es un placeholder para la implementación específica de impresión
            Cursor.Current = Cursors.WaitCursor

            Try
                Select Case wPuestoTrabajo.TipoImpresora
                    Case "TEC"
                        wsImprimirEtiquetasCajasImpresoraTec(vtDatosEtiquetasCajas, blPantalla, MOD_Nombre)
                    Case "ZEBRA"
                        wsImprimirEtiquetasCajasImpresoraZebra(vtDatosEtiquetasCajas, blPantalla, MOD_Nombre)
                    Case Else
                        wsImprimirEtiquetasCajasImpresoraNormal(vtDatosEtiquetasCajas, blPantalla, MOD_Nombre)
                End Select
            Finally
                Cursor.Current = Cursors.Default
            End Try
        End Sub

        Private Sub wsImprimirEtiquetasCajasImpresoraTec(vtDatosEtiquetasCajas As Object(,), blPantalla As Boolean, MOD_Nombre As String)
            ' Implementación para impresora TEC
            ' TODO: Implementar según las especificaciones de la impresora TEC
            MessageBox.Show("Impresión TEC no implementada en esta versión.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Sub wsImprimirEtiquetasCajasImpresoraZebra(vtDatosEtiquetasCajas As Object(,), blPantalla As Boolean, MOD_Nombre As String)
            ' Implementación para impresora Zebra
            ' TODO: Implementar según las especificaciones de la impresora Zebra
            MessageBox.Show("Impresión Zebra no implementada en esta versión.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Sub wsImprimirEtiquetasCajasImpresoraNormal(vtDatosEtiquetasCajas As Object(,), blPantalla As Boolean, MOD_Nombre As String)
            ' Implementación para impresora normal
            ' TODO: Implementar usando System.Drawing.Printing
            MessageBox.Show("Impresión normal no implementada en esta versión.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        '**************************************************************************************
        'Función:   wfValidarEstadoGrupo
        'Objetivo:  Validar el Estado de un Grupo
        '**************************************************************************************
        Public Function wfValidarEstadoGrupo(stDonde As String,
                                              Optional ilGrupo As Long = 0,
                                              Optional stCodigoEstado As String = "",
                                              Optional stDescripcionEstado As String = "",
                                              Optional stFinEstadoGrupo As String = "-1",
                                              Optional blMensaje As Boolean = True) As Boolean
            Dim blSeguir As Boolean = True

            Select Case stDonde
                Case "ASIGNAR_ARTICULO_TABLILLA", "ASOCIAR_UNIDADES_TRANSPORTE"
                    ' Si no se manda el estado hay que averiguarlo del Grupo
                    If String.IsNullOrEmpty(stCodigoEstado) Then
                        blSeguir = False

                        Try
                            Using conn As New SqlConnection(ConexionGestion)
                                conn.Open()
                                Dim sql As String = $"SELECT * FROM gacgrupo WITH (INDEX = gacgrupo_cgrcod, NOLOCK) 
                                                     LEFT JOIN gaestgru WITH (INDEX = gaestgru_estcod, NOLOCK) ON (cgrest = estcod) 
                                                     WHERE cgrcod = {ilGrupo}"

                                Using cmd As New SqlCommand(sql, conn)
                                    Using reader As SqlDataReader = cmd.ExecuteReader()
                                        If reader.Read() Then
                                            stCodigoEstado = reader("estcod").ToString()
                                            stDescripcionEstado = reader("estdes").ToString()
                                            stFinEstadoGrupo = reader("estfin").ToString()
                                            blSeguir = True
                                        Else
                                            If blMensaje Then wsMensaje($" No existe el Grupo {ilGrupo}", TipoMensaje.MENSAJE_Grave)
                                        End If
                                    End Using
                                End Using
                            End Using
                        Catch ex As Exception
                            If blMensaje Then wsMensaje($" Error al verificar grupo: {ex.Message}", TipoMensaje.MENSAJE_Grave)
                            Return False
                        End Try
                    End If

                    If blSeguir Then
                        If stFinEstadoGrupo = "0" Then
                            ' Sólo lo valida si está Iniciado
                            If stCodigoEstado = EstadoGrupo_Iniciado Then
                                Return True
                            Else
                                If blMensaje Then wsMensaje($" El Grupo está en Estado {stDescripcionEstado}, que no permite su modificación ", TipoMensaje.MENSAJE_Grave)
                            End If
                        Else
                            If blMensaje Then wsMensaje($" El Grupo {ilGrupo}, está en estado {stDescripcionEstado}, que no permite su modificación ", TipoMensaje.MENSAJE_Grave)
                        End If
                    End If
            End Select

            Return False
        End Function

        '**************************************************************************************
        'Función:   wfCrearAsignacion
        'Objetivo:  Crear la asociación del Bac a la Tablilla
        '**************************************************************************************
        Public Function wfCrearAsignacion(ilGrupo As Long, ilTablilla As Long, stBac As String) As Boolean
            Try
                Using conn As New SqlConnection(ConexionGestion)
                    conn.Open()

                    ' Comprobación de la existencia de la tablilla
                    Dim checkSql As String = $"SELECT * FROM GACTABLI WHERE CTAGRU={ilGrupo} AND CTATAB={ilTablilla}"
                    Using checkCmd As New SqlCommand(checkSql, conn)
                        Using reader As SqlDataReader = checkCmd.ExecuteReader()
                            If Not reader.HasRows Then
                                wsMensaje($" La tablilla {ilTablilla} del grupo {ilGrupo} no existe. No se puede crear la asignación", TipoMensaje.MENSAJE_Grave)
                                Return False
                            End If
                        End Using
                    End Using

                    ' Asignación del BAC a la tablilla usando stored procedure
                    Using cmd As New SqlCommand("dbo.AsignacionBacATablilla", conn)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.AddWithValue("@bac", stBac)
                        cmd.Parameters.AddWithValue("@grupo", ilGrupo)
                        cmd.Parameters.AddWithValue("@tablilla", ilTablilla)
                        cmd.Parameters.AddWithValue("@usuario", Usuario.Id)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using

                wsMensaje($" Asociada la tablilla {ilTablilla} del grupo {ilGrupo} al BAC {stBac}", TipoMensaje.MENSAJE_Informativo)
                Return True

            Catch ex As Exception
                wsMensaje($" Error de asignación del bac: {stBac} - {ex.Message}", TipoMensaje.MENSAJE_Grave)
                Return False
            End Try
        End Function

        '**************************************************************************************
        'Función:   wfValidarBAC
        'Objetivo:  Función para Validar la Existencia de un BAC en el maestro de BAC
        '**************************************************************************************
        Public Function wfValidarBAC(vtBAC As Object, Optional blMensaje As Boolean = True) As Boolean
            Try
                Using conn As New SqlConnection(ConexionGestion)
                    conn.Open()
                    Dim sql As String = "SELECT * FROM GAUBIBAC WHERE UBIBAC = @BAC"

                    Using cmd As New SqlCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@BAC", If(vtBAC, DBNull.Value))
                        Using reader As SqlDataReader = cmd.ExecuteReader()
                            If reader.HasRows Then
                                Return True
                            Else
                                If blMensaje Then wsMensaje($" No existe el BAC {vtBAC}", TipoMensaje.MENSAJE_Grave)
                                Return False
                            End If
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                If blMensaje Then wsMensaje($" Error al validar BAC: {ex.Message}", TipoMensaje.MENSAJE_Grave)
                Return False
            End Try
        End Function

        '**************************************************************************************
        'Función:   wfValidarEAN13
        'Objetivo:  Función para Validar la Existencia de un EAN13 o UPC12 en el maestro de Artículos
        '**************************************************************************************
        Public Function wfValidarEAN13(vtEAN13 As Object, Optional blMensaje As Boolean = True) As Boolean
            Try
                Using conn As New SqlConnection(ConexionGestion)
                    conn.Open()
                    Dim sql As String = "SELECT * FROM gaarticu WITH (INDEX = gaarticu_artean, NOLOCK) WHERE artean = @EAN13"

                    Using cmd As New SqlCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@EAN13", If(vtEAN13, DBNull.Value))
                        Using reader As SqlDataReader = cmd.ExecuteReader()
                            If reader.HasRows Then
                                Return True
                            Else
                                If blMensaje Then wsMensaje($"Error Ean13: {vtEAN13}", TipoMensaje.MENSAJE_Grave)
                                Return False
                            End If
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                If blMensaje Then wsMensaje($" Error al validar EAN13: {ex.Message}", TipoMensaje.MENSAJE_Grave)
                Return False
            End Try
        End Function

    End Module

End Namespace
