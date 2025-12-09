'*****************************************************************************************
'clGenericaRecordset.vb
'
' Clase genérica para manejo de Recordsets
' Converted from VB6 to VB.NET
' Esta clase proporciona funcionalidad similar al Recordset desconectado de ADO
'*****************************************************************************************

Imports System
Imports System.Data

Namespace Classes

    Public Class clGenericaRecordset
        Implements IDisposable

        Private _dataTable As DataTable
        Private _currentPosition As Integer = -1
        Private _disposed As Boolean = False

        ' Constructor
        Public Sub New()
            _dataTable = New DataTable()
            _currentPosition = -1
        End Sub

        ' Propiedad para obtener el número de registros
        Public ReadOnly Property RecordCount As Integer
            Get
                If _dataTable Is Nothing Then Return 0
                Return _dataTable.Rows.Count
            End Get
        End Property

        ' Propiedad EOF (End of File)
        Public ReadOnly Property EOF As Boolean
            Get
                Return _currentPosition >= RecordCount OrElse RecordCount = 0
            End Get
        End Property

        ' Propiedad BOF (Beginning of File)
        Public ReadOnly Property BOF As Boolean
            Get
                Return _currentPosition < 0 OrElse RecordCount = 0
            End Get
        End Property

        ' Obtener el DataTable interno
        Public ReadOnly Property DataTable As DataTable
            Get
                Return _dataTable
            End Get
        End Property

        ' Obtener el campo por índice
        Public ReadOnly Property Campo(index As Integer) As Object
            Get
                If _currentPosition >= 0 AndAlso _currentPosition < RecordCount Then
                    Return _dataTable.Rows(_currentPosition)(index)
                End If
                Return Nothing
            End Get
        End Property

        ' Obtener el campo por nombre
        Public ReadOnly Property Campo(name As String) As Object
            Get
                If _currentPosition >= 0 AndAlso _currentPosition < RecordCount Then
                    If _dataTable.Columns.Contains(name) Then
                        Return _dataTable.Rows(_currentPosition)(name)
                    End If
                End If
                Return Nothing
            End Get
        End Property

        ''' <summary>
        ''' Configura las columnas del recordset
        ''' Los argumentos se pasan como grupos de 3: nombre, tipo ADO, tamaño
        ''' </summary>
        Public Sub Configura_Clase(ParamArray args() As Object)
            _dataTable = New DataTable()
            _currentPosition = -1

            ' Procesar argumentos en grupos de 3 (nombre, tipo, tamaño)
            Dim i As Integer = 2 ' Saltar los primeros 2 parámetros (cursorType, lockType)
            While i + 2 < args.Length
                Dim nombre As String = args(i).ToString()
                Dim tipoAdo As Integer = CInt(args(i + 1))
                Dim tamaño As Integer = CInt(args(i + 2))

                Dim tipoColumna As Type = ConvertirTipoADO(tipoAdo)
                _dataTable.Columns.Add(nombre, tipoColumna)

                i += 3
            End While
        End Sub

        ' Convertir tipos ADO a tipos .NET
        Private Function ConvertirTipoADO(adoType As Integer) As Type
            ' Constantes ADO comunes
            Const adInteger As Integer = 3
            Const adSmallInt As Integer = 2
            Const adVarChar As Integer = 200
            Const adWChar As Integer = 130
            Const adDate As Integer = 7
            Const adDouble As Integer = 5
            Const adDecimal As Integer = 14
            Const adBoolean As Integer = 11

            Select Case adoType
                Case adInteger, adSmallInt
                    Return GetType(Integer)
                Case adVarChar, adWChar
                    Return GetType(String)
                Case adDate
                    Return GetType(Date)
                Case adDouble, adDecimal
                    Return GetType(Double)
                Case adBoolean
                    Return GetType(Boolean)
                Case Else
                    Return GetType(Object)
            End Select
        End Function

        ' Agregar un nuevo registro
        Public Sub Add(ParamArray values() As Object)
            If _dataTable Is Nothing Then Return

            Dim row As DataRow = _dataTable.NewRow()
            For i As Integer = 0 To Math.Min(values.Length - 1, _dataTable.Columns.Count - 1)
                row(i) = If(values(i), DBNull.Value)
            Next
            _dataTable.Rows.Add(row)

            ' Posicionarse en el nuevo registro
            _currentPosition = _dataTable.Rows.Count - 1
        End Sub

        ' Eliminar el registro actual
        Public Sub Delete()
            If _currentPosition >= 0 AndAlso _currentPosition < RecordCount Then
                _dataTable.Rows.RemoveAt(_currentPosition)
                If _currentPosition >= RecordCount Then
                    _currentPosition = RecordCount - 1
                End If
            End If
        End Sub

        ' Moverse al primer registro
        Public Sub MoveFirst()
            If RecordCount > 0 Then
                _currentPosition = 0
            Else
                _currentPosition = -1
            End If
        End Sub

        ' Moverse al último registro
        Public Sub MoveLast()
            If RecordCount > 0 Then
                _currentPosition = RecordCount - 1
            Else
                _currentPosition = -1
            End If
        End Sub

        ' Moverse al siguiente registro
        Public Sub MoveNext()
            If _currentPosition < RecordCount Then
                _currentPosition += 1
            End If
        End Sub

        ' Moverse al registro anterior
        Public Sub MovePrevious()
            If _currentPosition > 0 Then
                _currentPosition -= 1
            ElseIf _currentPosition = 0 Then
                _currentPosition = -1
            End If
        End Sub

        ' Limpiar todos los registros
        Public Sub Clear()
            _dataTable.Clear()
            _currentPosition = -1
        End Sub

        ' Implementación de IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    If _dataTable IsNot Nothing Then
                        _dataTable.Dispose()
                        _dataTable = Nothing
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
