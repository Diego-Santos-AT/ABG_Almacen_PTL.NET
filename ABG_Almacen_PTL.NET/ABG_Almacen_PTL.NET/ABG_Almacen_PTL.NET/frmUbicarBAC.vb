'***********************************************************************
'Nombre: frmUbicarBAC.vb
' Formulario para la ubicación de un BAC en una ubicación de PTL
' Converted from VB6 to VB.NET - Faithful line-by-line conversion
'
'Creación:      05/06/20
'
'Realización:   A.Esteban
'
'***********************************************************************

Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports ABG_Almacen_PTL.Modules
Imports ABG_Almacen_PTL.DataAccess

Public Class frmUbicarBAC
    Inherits Form

    ' ----- Constantes de Módulo (igual que VB6) -------------
    Private Const MOD_Nombre As String = "Extraer BAC"  ' VB6: MOD_Nombre = "Extraer BAC"
    Private Const CML_Salir As Integer = 990
    Private Const CML_Cancelar As Integer = 0
    Private Const LIS_ContenidoBAC As Integer = 1
    Private Const ColorRojo As Integer = &H8080FF
    Private Const ColorVerde As Integer = &H80FF80

    ' ----- Variables generales (igual que VB6) -------------
    Private ed As EntornoDeDatos
    Private iUbicacion As Integer
    ' Private CustomDataFilter As clsDataFilter ' VB6 compatibility (not implemented)

    ' Controles principales
    Private FrameLectura As Panel
    Private WithEvents txtLecturaCodigo As TextBox
    Private cmdSalir As Button      ' cmdAccion(990)
    Private cmdCancelar As Button   ' cmdAccion(0)

    ' Panel de opciones
    Private Picture1 As Panel
    Private oEstado(1) As RadioButton  ' Array igual que VB6

    ' Frame datos
    Private fraArticulo As Panel

    ' Labels - títulos
    Private lbTexto As Label
    Private Label8 As Label
    Private Label1 As Label
    Private Label2 As Label
    Private Label3 As Label
    Private Label4 As Label
    Private Label5 As Label
    Private Label6 As Label
    Private Label7 As Label
    Private Label9 As Label

    ' Labels - valores
    Private lbUbicacion As Label
    Private lbBAC As Label
    Private lbEstadoBAC As Label
    Private lbGrupo As Label
    Private lbTablilla As Label
    Private lbUds As Label
    Private lbPeso As Label
    Private lbVolumen As Label
    Private lbTipoCaja As Label
    Private lbNombreCaja As Label

    '*******************************************************************************

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        ' Propiedades del formulario - igual que VB6
        ' VB6: BackColor = &H00B06000& (BGR) = RGB(0, 96, 176)
        Me.BackColor = Color.FromArgb(0, 96, 176)
        Me.FormBorderStyle = FormBorderStyle.None  ' BorderStyle = 0 'None
        Me.Text = "Form1"                           ' Caption = "Form1"
        Me.ClientSize = New Size(273, 301)          ' 4095x4515 twips / 15
        Me.KeyPreview = True                        ' KeyPreview = -1 'True
        Me.ShowInTaskbar = False                    ' ShowInTaskbar = 0 'False
        Me.StartPosition = FormStartPosition.Manual

        ' FrameLectura (PictureBox en VB6)
        ' VB6: BackColor = &H00808080& = Gray
        FrameLectura = New Panel()
        FrameLectura.BackColor = Color.Gray
        FrameLectura.Location = New Point(0, 0)
        FrameLectura.Size = New Size(254, 301)      ' 3810x4515 twips / 15

        ' lbTexto - Título
        lbTexto = New Label()
        lbTexto.Text = "UBICAR BAC"
        lbTexto.Font = New Font("MS Sans Serif", 12, FontStyle.Bold)
        lbTexto.ForeColor = Color.White
        lbTexto.BackColor = Color.Gray
        lbTexto.TextAlign = ContentAlignment.MiddleCenter
        lbTexto.Location = New Point(5, 5)          ' 75/15, 75/15
        lbTexto.Size = New Size(245, 16)            ' 3675/15, 240/15

        ' txtLecturaCodigo
        txtLecturaCodigo = New TextBox()
        txtLecturaCodigo.Font = New Font("Arial", 14.25F)
        txtLecturaCodigo.BackColor = Color.White
        txtLecturaCodigo.Location = New Point(5, 30)    ' 75/15, 450/15
        txtLecturaCodigo.Size = New Size(180, 30)       ' 2700/15
        txtLecturaCodigo.MaxLength = 36

        ' cmdSalir (cmdAccion(990) en VB6)
        cmdSalir = New Button()
        cmdSalir.Text = "SALIR"
        cmdSalir.Location = New Point(190, 30)      ' 2850/15, 450/15
        cmdSalir.Size = New Size(61, 30)            ' 915/15, 450/15
        cmdSalir.TabStop = False
        AddHandler cmdSalir.Click, AddressOf cmdSalir_Click

        ' Label8 - Instrucción
        Label8 = New Label()
        Label8.Text = "Leer BAC o Ubicación de PTL"
        Label8.BackColor = SystemColors.Window
        Label8.TextAlign = ContentAlignment.MiddleCenter
        Label8.Location = New Point(5, 70)          ' 75/15, 1050/15
        Label8.Size = New Size(245, 21)             ' 3675/15, 315/15

        ' Picture1 - Panel de opciones
        ' VB6: BackColor = &H00C0C0C0& = Silver
        Picture1 = New Panel()
        Picture1.BackColor = Color.FromArgb(&HC0, &HC0, &HC0)
        Picture1.Location = New Point(5, 95)        ' 75/15, 1425/15
        Picture1.Size = New Size(245, 61)           ' 3675/15, 915/15

        ' oEstado(0) - Cerrar BAC
        oEstado(0) = New RadioButton()
        oEstado(0).Text = "Cerrar BAC"
        oEstado(0).BackColor = Color.FromArgb(&HC0, &HC0, &HC0)
        oEstado(0).Location = New Point(10, 5)      ' 150/15, 75/15
        oEstado(0).Size = New Size(121, 21)         ' 1815/15, 315/15
        AddHandler oEstado(0).Click, AddressOf oEstado_Click

        ' oEstado(1) - Abrir BAC
        oEstado(1) = New RadioButton()
        oEstado(1).Text = "Abrir BAC"
        oEstado(1).BackColor = Color.FromArgb(&HC0, &HC0, &HC0)
        oEstado(1).Location = New Point(10, 30)     ' 150/15, 450/15
        oEstado(1).Size = New Size(121, 21)
        oEstado(1).Checked = True                   ' Value = -1 'True
        AddHandler oEstado(1).Click, AddressOf oEstado_Click

        ' cmdCancelar (cmdAccion(0) en VB6)
        cmdCancelar = New Button()
        cmdCancelar.Text = "CANCELAR"
        cmdCancelar.Location = New Point(155, 15)   ' 2325/15, 225/15
        cmdCancelar.Size = New Size(66, 30)         ' 990/15, 450/15
        cmdCancelar.TabStop = False
        AddHandler cmdCancelar.Click, AddressOf cmdCancelar_Click

        Picture1.Controls.AddRange(New Control() {oEstado(0), oEstado(1), cmdCancelar})

        ' fraArticulo - Frame de datos
        fraArticulo = New Panel()
        fraArticulo.BackColor = Color.FromArgb(&HC0, &HC0, &HC0)
        fraArticulo.Location = New Point(5, 160)    ' 75/15, 2400/15
        fraArticulo.Size = New Size(245, 130)       ' 3675/15, 1950/15

        ' Labels de etiquetas
        Label2 = New Label()
        Label2.Text = "Ubicación"
        Label2.Location = New Point(5, 5)           ' 75/15, 75/15
        Label2.Size = New Size(51, 19)              ' 765/15, 285/15

        Label3 = New Label()
        Label3.Text = "BAC"
        Label3.Location = New Point(5, 30)          ' 75/15, 450/15
        Label3.Size = New Size(51, 19)

        Label1 = New Label()
        Label1.Text = "Grupo"
        Label1.Location = New Point(5, 55)          ' 75/15, 825/15
        Label1.Size = New Size(51, 19)

        Label4 = New Label()
        Label4.Text = "Tablilla"
        Label4.Location = New Point(100, 55)        ' 1500/15
        Label4.Size = New Size(41, 19)

        Label9 = New Label()
        Label9.Text = "Uds"
        Label9.Location = New Point(175, 55)        ' 2625/15
        Label9.Size = New Size(31, 19)

        Label5 = New Label()
        Label5.Text = "Peso"
        Label5.Location = New Point(5, 80)          ' 75/15, 1200/15
        Label5.Size = New Size(51, 19)

        Label6 = New Label()
        Label6.Text = "Volumen"
        Label6.Location = New Point(115, 80)        ' 1725/15
        Label6.Size = New Size(56, 19)

        Label7 = New Label()
        Label7.Text = "Tipo Caja"
        Label7.Location = New Point(5, 105)         ' 75/15, 1575/15
        Label7.Size = New Size(51, 19)

        ' Labels de valores
        lbUbicacion = New Label()
        lbUbicacion.BackColor = Color.White
        lbUbicacion.Font = New Font("MS Sans Serif", 9.75F, FontStyle.Bold)
        lbUbicacion.TextAlign = ContentAlignment.MiddleCenter
        lbUbicacion.Location = New Point(60, 5)     ' 900/15
        lbUbicacion.Size = New Size(177, 19)        ' 2655/15

        lbBAC = New Label()
        lbBAC.BackColor = Color.White
        lbBAC.Font = New Font("MS Sans Serif", 9.75F, FontStyle.Bold)
        lbBAC.TextAlign = ContentAlignment.MiddleCenter
        lbBAC.Location = New Point(60, 30)
        lbBAC.Size = New Size(102, 19)              ' 1530/15

        lbEstadoBAC = New Label()
        lbEstadoBAC.BackColor = Color.White
        lbEstadoBAC.TextAlign = ContentAlignment.MiddleCenter
        lbEstadoBAC.Location = New Point(166, 30)   ' 2490/15
        lbEstadoBAC.Size = New Size(71, 19)         ' 1065/15

        lbGrupo = New Label()
        lbGrupo.BackColor = Color.White
        lbGrupo.TextAlign = ContentAlignment.MiddleCenter
        lbGrupo.Location = New Point(60, 55)
        lbGrupo.Size = New Size(37, 19)             ' 555/15

        lbTablilla = New Label()
        lbTablilla.BackColor = Color.White
        lbTablilla.TextAlign = ContentAlignment.MiddleCenter
        lbTablilla.Location = New Point(145, 55)    ' 2175/15
        lbTablilla.Size = New Size(27, 19)          ' 405/15

        lbUds = New Label()
        lbUds.BackColor = Color.White
        lbUds.TextAlign = ContentAlignment.MiddleCenter
        lbUds.Location = New Point(210, 55)         ' 3150/15
        lbUds.Size = New Size(27, 19)               ' 405/15

        lbPeso = New Label()
        lbPeso.BackColor = Color.White
        lbPeso.TextAlign = ContentAlignment.MiddleCenter
        lbPeso.Location = New Point(60, 80)
        lbPeso.Size = New Size(52, 19)              ' 780/15

        lbVolumen = New Label()
        lbVolumen.BackColor = Color.White
        lbVolumen.TextAlign = ContentAlignment.MiddleCenter
        lbVolumen.Location = New Point(175, 80)     ' 2625/15
        lbVolumen.Size = New Size(62, 19)           ' 930/15

        lbTipoCaja = New Label()
        lbTipoCaja.BackColor = Color.White
        lbTipoCaja.TextAlign = ContentAlignment.MiddleCenter
        lbTipoCaja.Location = New Point(60, 105)
        lbTipoCaja.Size = New Size(37, 19)          ' 555/15

        lbNombreCaja = New Label()
        lbNombreCaja.BackColor = Color.White
        lbNombreCaja.TextAlign = ContentAlignment.MiddleCenter
        lbNombreCaja.Location = New Point(101, 105) ' 1515/15
        lbNombreCaja.Size = New Size(136, 19)       ' 2040/15

        ' Agregar labels al frame de datos
        fraArticulo.Controls.AddRange(New Control() {
            Label2, Label3, Label1, Label4, Label9, Label5, Label6, Label7,
            lbUbicacion, lbBAC, lbEstadoBAC, lbGrupo, lbTablilla, lbUds,
            lbPeso, lbVolumen, lbTipoCaja, lbNombreCaja
        })

        ' Agregar controles al FrameLectura
        FrameLectura.Controls.AddRange(New Control() {
            lbTexto, txtLecturaCodigo, cmdSalir, Label8, Picture1, fraArticulo
        })

        Me.Controls.Add(FrameLectura)

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    '*******************************************************************************

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Cursor = Cursors.WaitCursor

        Me.Top = 0
        Me.Left = 0
        Me.Width = 254    ' 3805/15
        Me.Height = 302   ' 4525/15

        Me.Text = MOD_Nombre

        ' Inicia el entorno de datos
        ed = New EntornoDeDatos()
        Try
            ed.Open()
        Catch ex As Exception
            wsMensaje("Error de conexión: " & ex.Message, TipoMensaje.MENSAJE_Grave)
        End Try

        FrameLectura.Left = 0

        Cursor = Cursors.Default
    End Sub

    Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                Me.Close()
        End Select
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If ed IsNot Nothing Then
            ed.Dispose()
            ed = Nothing
        End If
    End Sub

    '----------------------------------------------------------------------------------

    Private Sub Salir()
        Me.Close()
    End Sub

    Private Sub cmdSalir_Click(sender As Object, e As EventArgs)
        Salir()
    End Sub

    Private Sub cmdCancelar_Click(sender As Object, e As EventArgs)
        Cancelar()
        txtLecturaCodigo.Focus()
    End Sub

    Private Sub oEstado_Click(sender As Object, e As EventArgs)
        txtLecturaCodigo.Focus()
    End Sub

    Private Sub txtLecturaCodigo_GotFocus(sender As Object, e As EventArgs) Handles txtLecturaCodigo.GotFocus
        txtLecturaCodigo.BackColor = Color.FromArgb(&HC0, &HFF, &HC0)
    End Sub

    Private Sub txtLecturaCodigo_LostFocus(sender As Object, e As EventArgs) Handles txtLecturaCodigo.LostFocus
        txtLecturaCodigo.BackColor = SystemColors.Window
    End Sub

    Private Sub txtLecturaCodigo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtLecturaCodigo.KeyDown
        If e.KeyCode = Keys.Return Then
            ' Inicializa la visualización
            ' RefrescarDatos(True)

            Select Case txtLecturaCodigo.Text.Length
                Case 12 ' Unidad de transporte / Ubicación
                    ' Comprobar si la lectura es un BAC
                    If fValidarBAC(txtLecturaCodigo.Text, False) = False Then
                        ' Comprobar si la lectura es una ubicación
                        If fValidarUbicacion(txtLecturaCodigo.Text, False) = False Then
                            ' No existe la ubicación / BAC
                            wsMensaje(" No se ha encontrado Ubicación o BAC", TipoMensaje.MENSAJE_Grave)
                        End If
                    End If
            End Select

            txtLecturaCodigo.Text = ""
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Function fValidarBAC(ByVal stBAC As String, Optional ByVal blMensaje As Boolean = True) As Boolean
        Dim bCalculoPeso As Boolean
        Dim bCalculoVolumen As Boolean
        Dim tEstado As Integer

        fValidarBAC = False

        Try
            Dim dtBAC As DataTable = ed.DameDatosBACdePTL(stBAC)

            If dtBAC.Rows.Count > 0 Then
                fValidarBAC = True
                Dim row As DataRow = dtBAC.Rows(0)

                Dim unipes As Double = If(Not IsDBNull(row("unipes")), CDbl(row("unipes")), 0)
                Dim unipma As Double = If(Not IsDBNull(row("unipma")), CDbl(row("unipma")), 0)
                Dim univol As Double = If(Not IsDBNull(row("univol")), CDbl(row("univol")), 0)
                Dim univma As Double = If(Not IsDBNull(row("univma")), CDbl(row("univma")), 0)
                Dim uninum As Integer = If(Not IsDBNull(row("uninum")), CInt(row("uninum")), 0)
                Dim unicod As String = If(Not IsDBNull(row("unicod")), row("unicod").ToString(), "")
                Dim uniest As Integer = If(Not IsDBNull(row("uniest")), CInt(row("uniest")), 0)
                Dim unigru As Integer = If(Not IsDBNull(row("unigru")), CInt(row("unigru")), 0)
                Dim unitab As Integer = If(Not IsDBNull(row("unitab")), CInt(row("unitab")), 0)
                Dim unicaj As String = If(Not IsDBNull(row("unicaj")), row("unicaj").ToString(), "")
                Dim tipdes As String = If(Not IsDBNull(row("tipdes")), row("tipdes").ToString(), "")

                bCalculoPeso = unipes > unipma
                bCalculoVolumen = univol > univma

                ' Se muestran los datos
                If IsDBNull(row("ubicod")) Then
                    RefrescarDatos(False, 0, 0, 0, 0, 0, unicod, uniest, unigru, unitab, unipes, univol, unicaj, tipdes, bCalculoPeso, bCalculoVolumen)
                Else
                    Dim ubicod As Integer = CInt(row("ubicod"))
                    Dim ubialm As Integer = If(Not IsDBNull(row("ubialm")), CInt(row("ubialm")), 0)
                    Dim ubiblo As Integer = If(Not IsDBNull(row("ubiblo")), CInt(row("ubiblo")), 0)
                    Dim ubifil As Integer = If(Not IsDBNull(row("ubifil")), CInt(row("ubifil")), 0)
                    Dim ubialt As Integer = If(Not IsDBNull(row("ubialt")), CInt(row("ubialt")), 0)
                    RefrescarDatos(False, ubicod, ubialm, ubiblo, ubifil, ubialt, unicod, uniest, unigru, unitab, unipes, univol, unicaj, tipdes, bCalculoPeso, bCalculoVolumen)
                End If

                If uninum > 0 Then
                    wsMensaje(" El BAC ya se encuentra ubicado ", TipoMensaje.MENSAJE_Grave)
                    RefrescarDatos(True)
                    iUbicacion = 0
                Else
                    If iUbicacion > 0 Then
                        ' Se ha leido la ubicación anteriormente. Se procede a ubicar el BAC
                        If UbicarBAC(unicod, iUbicacion, uniest, oEstado(0).Checked) Then
                            wsMensaje(" Se ha ubicado el BAC: " & unicod & " en la ubicación de PTL " & iUbicacion.ToString(), TipoMensaje.MENSAJE_Exclamacion)
                            iUbicacion = 0  ' Reiniciamos la ubicación
                            RefrescarDatos(True)
                        End If
                    Else
                        ' Se queda pendiente de la lectura de la ubicación o de otro BAC
                    End If
                End If
            Else
                ' No se ha encontrado el BAC. Se comprueba si existe la definición en GAUBIBAC
                Dim dtConsulta As DataTable = ed.ConsultaBACdePTL(stBAC)

                If dtConsulta.Rows.Count > 0 Then
                    ' Se ha encontrado el BAC
                    fValidarBAC = True
                    Dim ubibac As String = If(Not IsDBNull(dtConsulta.Rows(0)("ubibac")), dtConsulta.Rows(0)("ubibac").ToString(), "")
                    RefrescarDatos(False, 0, 0, 0, 0, 0, ubibac, 0, 0, 0, 0, 0, "", "", False, False)
                    If iUbicacion > 0 Then
                        ' Se ha leido la ubicación anteriormente. Se procede a ubicar el BAC
                        If UbicarBAC(ubibac, iUbicacion, 0, oEstado(0).Checked) Then
                            wsMensaje(" Se ha ubicado el BAC: " & ubibac & " en la ubicación de PTL " & iUbicacion.ToString(), TipoMensaje.MENSAJE_Exclamacion)
                            iUbicacion = 0  ' Reiniciamos la ubicación
                            RefrescarDatos(True)
                        End If
                    Else
                        ' Se queda pendiente de la lectura de la ubicación o de otro BAC
                    End If
                Else
                    If blMensaje Then wsMensaje(" No existe el BAC ", TipoMensaje.MENSAJE_Grave)
                End If
            End If
        Catch ex As Exception
            If blMensaje Then wsMensaje($" Error: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        Return fValidarBAC
    End Function

    Private Function fValidarUbicacion(ByVal stUbicacion As String, Optional ByVal blMensaje As Boolean = True) As Boolean
        Dim iALF As Integer
        Dim iALM As Integer
        Dim iBLO As Integer
        Dim iFIL As Integer
        Dim iALT As Integer

        fValidarUbicacion = False

        iALF = 2
        iALM = CInt(Val(stUbicacion.Substring(0, 3)))
        iBLO = CInt(Val(stUbicacion.Substring(3, 3)))
        iFIL = CInt(Val(stUbicacion.Substring(6, 3)))
        iALT = CInt(Val(stUbicacion.Substring(9, 3)))

        Try
            Dim dtUbicacion As DataTable = ed.DameDatosUbicacionPTL(iALF, iALM, iBLO, iFIL, iALT)

            If dtUbicacion.Rows.Count > 0 Then
                fValidarUbicacion = True
                Dim row As DataRow = dtUbicacion.Rows(0)
                Dim ubicod As Integer = If(Not IsDBNull(row("ubicod")), CInt(row("ubicod")), 0)
                iUbicacion = ubicod

                ' Si existe comprueba si tiene un BAC asociado
                If IsDBNull(row("unicod")) Then
                    lbUbicacion.Text = $"({ubicod}) {iALM:000}.{iBLO:000}.{iFIL:000}.{iALT:000}"
                    ' Si se ha leido el BAC anteriormente. Se procede a ubicar el BAC
                    If lbBAC.Text <> "" Then
                        Dim estadoBAC As Integer = If(lbEstadoBAC.Text = "ABIERTO", 0, 1)
                        If UbicarBAC(lbBAC.Text, iUbicacion, estadoBAC, oEstado(0).Checked) Then
                            wsMensaje(" Se ha ubicado el BAC: " & lbBAC.Text & " en la ubicación de PTL " & iUbicacion.ToString(), TipoMensaje.MENSAJE_Exclamacion)
                            iUbicacion = 0  ' Reiniciamos la ubicación
                            RefrescarDatos(True)
                        End If
                    Else
                        ' Se queda pendiente de la lectura del BAC o de otra ubicación
                    End If
                Else
                    wsMensaje(" La Ubicación ya tiene asociado un BAC ", TipoMensaje.MENSAJE_Grave)
                    iUbicacion = 0
                End If
            Else
                If blMensaje Then wsMensaje(" No existe la Unidad de Transporte ", TipoMensaje.MENSAJE_Grave)
                lbUbicacion.Text = ""
                iUbicacion = 0
            End If
        Catch ex As Exception
            If blMensaje Then wsMensaje($" Error: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        Return fValidarUbicacion
    End Function

    Private Sub RefrescarDatos(sEnBlanco As Boolean,
                               Optional sCodUbicacion As Integer = 0,
                               Optional sALM As Integer = 0,
                               Optional sBLO As Integer = 0,
                               Optional sFIL As Integer = 0,
                               Optional sALT As Integer = 0,
                               Optional sBAC As String = "",
                               Optional sEstadoBAC As Integer = 0,
                               Optional sGrupo As Integer = 0,
                               Optional sTablilla As Integer = 0,
                               Optional sPeso As Double = 0,
                               Optional sVolumen As Double = 0,
                               Optional sTipoCaja As String = "",
                               Optional sNombreCaja As String = "",
                               Optional bPeso As Boolean = False,
                               Optional bVolumen As Boolean = False)

        If sEnBlanco = True Then
            ' Inicia la visualización
            lbUbicacion.Text = ""
            lbBAC.Text = ""
            lbEstadoBAC.Text = ""
            lbEstadoBAC.BackColor = Color.White

            lbGrupo.Text = ""
            lbTablilla.Text = ""
            lbUds.Text = ""

            lbPeso.Text = ""
            lbPeso.BackColor = Color.White

            lbVolumen.Text = ""
            lbVolumen.BackColor = Color.White

            lbTipoCaja.Text = ""
            lbNombreCaja.Text = ""
        Else
            If sCodUbicacion = 0 Then
                lbUbicacion.Text = "SIN UBICACION"
                lbUbicacion.Text = "-------------"
            Else
                lbUbicacion.Text = $"({sCodUbicacion}) {sALM:000}.{sBLO:000}.{sFIL:000}.{sALT:000}"
            End If

            lbBAC.Text = sBAC
            lbEstadoBAC.Text = If(sEstadoBAC = 0, "ABIERTO", "CERRADO")
            If lbEstadoBAC.Text = "CERRADO" Then
                lbEstadoBAC.BackColor = Color.FromArgb((ColorVerde >> 16) And &HFF, (ColorVerde >> 8) And &HFF, ColorVerde And &HFF)
            End If

            lbGrupo.Text = sGrupo.ToString()
            lbTablilla.Text = sTablilla.ToString()
            lbUds.Text = "0"

            lbPeso.Text = sPeso.ToString("#0.000")
            If bPeso Then
                lbPeso.BackColor = Color.FromArgb((ColorRojo >> 16) And &HFF, (ColorRojo >> 8) And &HFF, ColorRojo And &HFF)
            End If

            lbVolumen.Text = sVolumen.ToString("#0.000")
            If bVolumen Then
                lbVolumen.BackColor = Color.FromArgb((ColorRojo >> 16) And &HFF, (ColorRojo >> 8) And &HFF, ColorRojo And &HFF)
            End If

            lbTipoCaja.Text = sTipoCaja
            lbNombreCaja.Text = sNombreCaja
        End If
    End Sub

    Private Sub Cancelar()
        RefrescarDatos(True)
        iUbicacion = 0
    End Sub

    Private Function UbicarBAC(tBac As String, tUbicacion As Integer, tEstado As Integer, tEstadoFinal As Boolean) As Boolean
        ' Ubicación del BAC
        Dim Retorno As Integer = 0
        Dim msgSalida As String = ""
        Dim nEstado As Integer

        UbicarBAC = False

        Try
            ed.UbicarBACenPTL(tBac, tUbicacion, Usuario.Id, Retorno, msgSalida)

            If Retorno = 0 Then
                UbicarBAC = True
                If (tEstado = 0) = tEstadoFinal Then
                    If tEstadoFinal Then nEstado = 1 Else nEstado = 0
                    ' Cambiar estado de BAC
                    If CambiarEstadoBAC(tBac, nEstado) Then
                        lbEstadoBAC.Text = If(tEstado = 0, "ABIERTO", "CERRADO")
                        If lbEstadoBAC.Text = "CERRADO" Then
                            lbEstadoBAC.BackColor = Color.FromArgb((ColorVerde >> 16) And &HFF, (ColorVerde >> 8) And &HFF, ColorVerde And &HFF)
                        Else
                            lbEstadoBAC.BackColor = Color.White
                        End If
                    End If
                End If
            Else
                wsMensaje(" No se ha podido ubicar el BAC en la estanteria de PTL. " & msgSalida, TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            wsMensaje($" Error al ubicar BAC: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        Return UbicarBAC
    End Function

    Private Function CambiarEstadoBAC(tBac As String, tEstado As Integer) As Boolean
        ' Cambio de estado de BAC
        Dim Retorno As Integer = 0
        Dim msgSalida As String = ""

        CambiarEstadoBAC = False

        Try
            ed.CambiaEstadoBACdePTL(tBac, tEstado, Usuario.Id, Retorno, msgSalida)

            If Retorno = 0 Then
                CambiarEstadoBAC = True
            Else
                wsMensaje(" No se ha podido cambiar el estado al BAC " & msgSalida, TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            wsMensaje($" Error al cambiar estado: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        Return CambiarEstadoBAC
    End Function

End Class
