'*****************************************************************************************
'Profile.vb
'
' Módulo genérico para lectura/escritura de archivos INI y registro de Windows
' Converted from VB6 to VB.NET
'=========================================================================================

Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Text

Namespace Modules

    Public Module Profile

        ' ---- Ruta y clave de registro para Archivos de programa ---------------------
        Public Const ProgramasDir As String = "SOFTWARE\Microsoft\Windows\CurrentVersion"
        Public Const ClaveRegistroProgramas As String = "ProgramFilesDir"
        '------------------------------------------------------------------------------

        ' ---- Ruta y clave de registro para Ruta DSN ---------------------------------
        Public Const DSNDir As String = "SOFTWARE\ODBC\ODBC.INI\ODBC File DSN"
        Public Const ClaveRegistroDSN As String = "DefaultDSNDir"
        '------------------------------------------------------------------------------

        ' Tipos ROOT de clave del Registro...
        Public Const HKEY_CURRENT_USER As Integer = &H80000001
        Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
        Public Const HKEY_USERS As Integer = &H80000003

        ' Win32 API Declarations for INI file operations
        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Function GetPrivateProfileString(
            lpApplicationName As String,
            lpKeyName As String,
            lpDefault As String,
            lpReturnedString As StringBuilder,
            nSize As Integer,
            lpFileName As String) As Integer
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Function WritePrivateProfileString(
            lpApplicationName As String,
            lpKeyName As String,
            lpString As String,
            lpFileName As String) As Integer
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Private Function GetPrivateProfileSection(
            lpAppName As String,
            lpReturnedString As StringBuilder,
            nSize As Integer,
            lpFileName As String) As Integer
        End Function

        ''' <summary>
        ''' Lee un valor de un archivo INI
        ''' </summary>
        ''' <param name="lpFileName">Ruta del archivo INI</param>
        ''' <param name="lpAppName">Nombre de la sección</param>
        ''' <param name="lpKeyName">Nombre de la clave</param>
        ''' <param name="vDefault">Valor por defecto si no se encuentra</param>
        ''' <returns>El valor leído o el valor por defecto</returns>
        Public Function LeerIni(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault As String = "") As String
            Dim sRetVal As New StringBuilder(255)
            Dim lTmp As Integer

            lTmp = GetPrivateProfileString(lpAppName, lpKeyName, vDefault, sRetVal, sRetVal.Capacity, lpFileName)

            If lTmp = 0 Then
                Return vDefault
            Else
                Return sRetVal.ToString(0, lTmp)
            End If
        End Function

        ''' <summary>
        ''' Guarda un valor en un archivo INI
        ''' </summary>
        ''' <param name="lpFileName">Ruta del archivo INI</param>
        ''' <param name="lpAppName">Nombre de la sección</param>
        ''' <param name="lpKeyName">Nombre de la clave</param>
        ''' <param name="lpString">Valor a guardar</param>
        Public Sub GuardarIni(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
            WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
        End Sub

        ''' <summary>
        ''' Lee una sección completa de un archivo INI
        ''' </summary>
        ''' <param name="lpFileName">Nombre del fichero INI</param>
        ''' <param name="lpAppName">Nombre de la sección a leer</param>
        ''' <returns>Una colección con el Valor y el contenido</returns>
        Public Function LeerSeccionINI(lpFileName As String, lpAppName As String) As System.Collections.Generic.Dictionary(Of String, String)
            Dim tContenidos As New System.Collections.Generic.Dictionary(Of String, String)
            Dim nSize As Integer
            Dim sBuffer As New StringBuilder(32767)

            nSize = GetPrivateProfileSection(lpAppName, sBuffer, sBuffer.Capacity, lpFileName)

            If nSize > 0 Then
                Dim strBuffer As String = sBuffer.ToString(0, nSize)

                ' Cada una de las entradas está separada por un Chr(0)
                Dim entries() As String = strBuffer.Split(New Char() {Chr(0)}, StringSplitOptions.RemoveEmptyEntries)

                For Each entry As String In entries
                    Dim j As Integer = entry.IndexOf("="c)
                    If j > 0 Then
                        Dim sClave As String = entry.Substring(0, j).Trim()
                        Dim sValor As String = If(j < entry.Length - 1, entry.Substring(j + 1).Trim(), "")
                        If Not tContenidos.ContainsKey(sClave) Then
                            tContenidos.Add(sClave, sValor)
                        End If
                    End If
                Next
            End If

            Return tContenidos
        End Function

        ''' <summary>
        ''' Lee una clave del registro de Windows
        ''' </summary>
        ''' <param name="KeyRoot">Raíz del registro (HKEY_LOCAL_MACHINE, etc.)</param>
        ''' <param name="KeyName">Ruta de la clave</param>
        ''' <param name="SubKeyRef">Nombre del valor</param>
        ''' <param name="KeyVal">Variable donde se almacenará el valor leído</param>
        ''' <returns>True si se leyó correctamente, False en caso contrario</returns>
        Public Function LeerClave(KeyRoot As Integer, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
            Try
                Dim baseKey As RegistryKey = Nothing

                Select Case KeyRoot
                    Case HKEY_LOCAL_MACHINE
                        baseKey = Registry.LocalMachine
                    Case HKEY_CURRENT_USER
                        baseKey = Registry.CurrentUser
                    Case HKEY_USERS
                        baseKey = Registry.Users
                    Case Else
                        Return False
                End Select

                Using subKey As RegistryKey = baseKey.OpenSubKey(KeyName, False)
                    If subKey Is Nothing Then
                        KeyVal = ""
                        Return False
                    End If

                    Dim value As Object = subKey.GetValue(SubKeyRef)
                    If value Is Nothing Then
                        KeyVal = ""
                        Return False
                    End If

                    KeyVal = value.ToString()
                    Return True
                End Using

            Catch ex As Exception
                KeyVal = ""
                Return False
            End Try
        End Function

    End Module

End Namespace
