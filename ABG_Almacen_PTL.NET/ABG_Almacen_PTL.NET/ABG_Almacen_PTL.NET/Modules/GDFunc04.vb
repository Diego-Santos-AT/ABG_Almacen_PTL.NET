'*****************************************************************************************
'GDFunc04.vb
'
' Módulo de funciones de utilidad
' Converted from VB6 to VB.NET
'
'*****************************************************************************************

Imports System
Imports System.Data
Imports Microsoft.Data.SqlClient
Imports System.Windows.Forms
Imports System.IO

Namespace Modules

    ' -- Tipo de Datos para contener registros seleccionados
    Public Structure Registro_Seleccionado
        Public iCodigo As Long      ' --- CODIGO
        Public sDescripcion As String  ' --- DESCRIPCION
    End Structure

    ' -- Tipo de Datos para contener las ubicaciones por defecto de empresa desglosada
    Public Structure Desglose_Ubicacion
        Public Almacen_Fisico As Integer
        Public Almacen_Logico As Integer
        Public Bloque As Integer
        Public Fila As Integer
        Public Altura As Integer
    End Structure

    Public Module GDFunc04

        '**************************************************************************************
        'Función: DesplazaRegistro2
        '
        'Función para simular la navegación sobre un recordset con los Movimientos
        ' tradicionales de Primero, Anterior, Siguiente y Ultimo
        '
        ' NOTA: Esta función utiliza nombres de tabla y columna pasados como parámetros.
        ' Estos valores provienen del código interno de la aplicación, no de entrada de usuario.
        ' Se validan los identificadores para prevenir inyección SQL.
        '**************************************************************************************
        Public Function DesplazaRegistro2(tabla As String, Campo As String, CodigoActual As String,
                                          TipoDesplazamiento As String, Optional condicion As String = "") As String
            ' Validar identificadores para prevenir inyección SQL
            If Not ValidarIdentificadorSQL(tabla) OrElse Not ValidarIdentificadorSQL(Campo) Then
                MessageBox.Show("Error: nombre de tabla o campo inválido", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return CodigoActual
            End If

            Dim Cond1 As String = ""

            If Not String.IsNullOrEmpty(condicion) Then
                Cond1 = " WHERE " & condicion
            End If

            Dim sql As String = ""
            Dim codigoParam As Long = 0
            Long.TryParse(CodigoActual, codigoParam)

            ' Usar el nombre de tabla/columna validado y parametrizar el valor
            Select Case TipoDesplazamiento
                Case "P"    'Primero
                    sql = $"SELECT MIN([{Campo}]) as Registro FROM [{tabla}]{Cond1}"

                Case "A"    'Anterior
                    If String.IsNullOrEmpty(Cond1) Then
                        Cond1 = $" WHERE [{Campo}] < @Codigo"
                    Else
                        Cond1 &= $" AND [{Campo}] < @Codigo"
                    End If
                    sql = $"SELECT MAX([{Campo}]) as Registro FROM [{tabla}]{Cond1}"

                Case "S"    'Siguiente
                    If String.IsNullOrEmpty(Cond1) Then
                        Cond1 = $" WHERE [{Campo}] > @Codigo"
                    Else
                        Cond1 &= $" AND [{Campo}] > @Codigo"
                    End If
                    sql = $"SELECT MIN([{Campo}]) as Registro FROM [{tabla}]{Cond1}"

                Case "U"    'Ultimo
                    sql = $"SELECT MAX([{Campo}]) as Registro FROM [{tabla}]{Cond1}"
            End Select

            Return EjecutaConsulta(sql, CodigoActual, codigoParam)
        End Function

        ''' <summary>
        ''' Valida que un identificador SQL (tabla o columna) sea seguro
        ''' </summary>
        Private Function ValidarIdentificadorSQL(identificador As String) As Boolean
            If String.IsNullOrEmpty(identificador) Then Return False
            ' Permitir solo caracteres alfanuméricos y guion bajo
            Return System.Text.RegularExpressions.Regex.IsMatch(identificador, "^[a-zA-Z_][a-zA-Z0-9_]*$")
        End Function

        Private Function EjecutaConsulta(sql As String, CodAct As String, Optional codigoParam As Long = 0) As String
            Try
                Cursor.Current = Cursors.WaitCursor

                Using conn As New SqlConnection(ConexionGestion)
                    conn.Open()
                    Using cmd As New SqlCommand(sql, conn)
                        ' Agregar parámetro si el SQL lo incluye
                        If sql.Contains("@Codigo") Then
                            cmd.Parameters.AddWithValue("@Codigo", codigoParam)
                        End If
                        Dim result As Object = cmd.ExecuteScalar()
                        If result Is Nothing OrElse IsDBNull(result) Then
                            Return CodAct
                        Else
                            Return result.ToString()
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error número: {ex.HResult} : {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return CodAct
            Finally
                Cursor.Current = Cursors.Default
            End Try
        End Function

        '- Función para truncar un numero decimal con el numero de decimales que se indique
        Public Function Truncar(Numero As Double, Numero_Decimales As Long) As Double
            Numero = Math.Round(Numero, 5)
            Dim factor As Double = Math.Pow(10, Numero_Decimales)
            Return Math.Truncate(Numero * factor) / factor
        End Function

        Public Function CambiaComaPorPunto(Precio As Object) As Object
            If Precio Is Nothing Then Return Precio
            Return Precio.ToString().Replace(",", ".")
        End Function

        '**************************************************************************************
        'Función: Exportar
        '
        'Función para exportar un DataTable a un fichero delimitado
        '**************************************************************************************
        Public Sub Exportar(dt As DataTable, Fichero As String, Optional Separador As String = "|")
            If String.IsNullOrEmpty(Fichero) OrElse dt Is Nothing Then Return

            Try
                Using writer As New StreamWriter(Fichero, False, System.Text.Encoding.UTF8)
                    ' Escribir cabeceras
                    Dim headers As String = String.Join(Separador, dt.Columns.Cast(Of DataColumn)().Select(Function(c) c.ColumnName))
                    writer.WriteLine(headers)

                    ' Escribir datos
                    For Each row As DataRow In dt.Rows
                        Dim values As String = String.Join(Separador, row.ItemArray.Select(Function(v) If(v Is Nothing OrElse IsDBNull(v), "", v.ToString())))
                        writer.WriteLine(values)
                    Next
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error al exportar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ' -- Procedimiento para Borrar los ficheros de un directorio según un patrón --
        Public Sub Limpiar_Temporales(Ruta As String, Patron As String)
            Try
                If Directory.Exists(Ruta) Then
                    Dim files() As String = Directory.GetFiles(Ruta, Patron)
                    For Each file As String In files
                        IO.File.Delete(file)
                    Next
                End If
            Catch ex As Exception
                ' Ignorar errores al limpiar temporales
            End Try
        End Sub

        ' -- Función que crear el Lote aplicado a una etiqueta ----
        Public Function Dame_Lote(Fecha_Hoy As Date) As String
            Dim lote As String = Dame_Fecha_Juliana(Fecha_Hoy)
            ' --- Le damos formato al lote con 8 Dígitos de Longitud ---------
            Return lote.PadLeft(8, "0"c)
        End Function

        '---- Procedimiento que calcula la Fecha Juliana a partir de una fecha determinada --
        Public Function Dame_Fecha_Juliana(Fecha As Date) As String
            Dim dia As Integer = Fecha.Day
            Dim mes As Integer = Fecha.Month
            Dim año As Integer = Fecha.Year

            ' Calculamos el número de días transcurridos según el período juliano hasta que comienza el Año
            Dim Z As Double = (4712 + año) * 365.25
            If Z = Math.Truncate(Z) Then Z = Z - 1 Else Z = Math.Truncate(Z)

            If (año <= 1583 OrElse (año = 1582 AndAlso (mes > 10 OrElse (mes = 10 AndAlso dia >= 15)))) AndAlso año <= 1700 Then Z = Z - 10
            If 1701 <= año AndAlso año <= 1800 Then Z = Z - 11
            If 1801 <= año AndAlso año <= 1900 Then Z = Z - 12
            If 1901 <= año AndAlso año <= 2100 Then Z = Z - 13
            If 2101 <= año AndAlso año <= 2200 Then Z = Z - 14

            ' Calculamos el número de días de los meses anteriores a Mes
            Dim B As Integer = Fecha.DayOfYear - dia

            ' Calculamos el día juliano
            Z = Z + B + dia

            Return CStr(CLng(Z))
        End Function

        ' ---- Función que Devuelve la Fecha y Hora del Sistema ---------------------
        Public Function Dame_FechaHora_Sistema() As Date
            Try
                Using conn As New SqlConnection(ConexionGestion)
                    conn.Open()
                    Using cmd As New SqlCommand("SELECT GETDATE() AS Hoy", conn)
                        Dim result As Object = cmd.ExecuteScalar()
                        If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                            Return CDate(result)
                        End If
                    End Using
                End Using
            Catch ex As Exception
                ' Si falla, devolver fecha local
            End Try

            Return DateTime.Now
        End Function

        ' --- Función para obtener la ruta completa de ubicación del fichero ABG.dsn --------
        Public Function Dame_DSN() As String
            Dim sAuxDSN As String = ""

            If Not LeerClave(HKEY_LOCAL_MACHINE, DSNDir, ClaveRegistroDSN, sAuxDSN) Then
                ' No se encontró el directorio de programas de Windows
                sAuxDSN = "C:\Archivos de programa\Archivos comunes\ODBC\Data Sources"
            End If

            RutaDSN = sAuxDSN & "\"
            Return RutaDSN & FicheroDSN
        End Function

        ' ---- Función para Crear el SSCC único para Una etiqueta
        Public Function Dame_SSCC(Empresa_Fabricante As Integer, Ean_Empresa_Fabricante As String,
                                   Numerador As Long, Optional IncrementoSerie As Integer = 0) As String

            ' -- Dígito de Extensión (Fijo a 3) --------------
            Dim Digito_Extension As String = "3"

            ' -- Código Ean de la Empresa --------------------
            Dim Ean_Empresa As String = Ean_Empresa_Fabricante

            ' -- Código de la serie (puede ser el código de empresa o el código de Fabricante asignado)
            Dim Codigo_Serie As String = (Empresa_Fabricante + IncrementoSerie).ToString("00")

            ' -- Si la empresa tiene más de dos dígitos toma los dos últimos
            If Codigo_Serie.Length > 2 Then
                Codigo_Serie = Codigo_Serie.Substring(Codigo_Serie.Length - 2, 2)
            End If

            ' --- Numerador único de bultos ----------------------
            Dim Numerador_Unico_Bultos As String = Numerador.ToString("0000000")

            ' --- Dígito de Control (calculado) ------------------
            Dim SSCC_Sin_DigitoControl As String = Digito_Extension & Ean_Empresa & Codigo_Serie & Numerador_Unico_Bultos

            Dim Suma_Parcial As Integer = 0
            For i As Integer = 0 To 16
                Dim digito As Integer = CInt(SSCC_Sin_DigitoControl.Substring(i, 1))
                Dim multiplicador As Integer = If((i Mod 2) = 0, 3, 1)
                Suma_Parcial += digito * multiplicador
            Next

            Dim Multiplo_Diez As Long = CLng(Math.Ceiling(Suma_Parcial / 10.0) * 10)
            Dim Digito_Control As Integer = CInt(Multiplo_Diez - Suma_Parcial)

            ' ----- CÓDIGO SSCC ----------------------------------
            Return Digito_Extension & Ean_Empresa & Codigo_Serie & Numerador_Unico_Bultos & Digito_Control.ToString()
        End Function

    End Module

End Namespace
