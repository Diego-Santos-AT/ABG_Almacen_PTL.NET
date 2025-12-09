'***********************************************************************
'Nombre: frmConsultaPTL.vb
' Formulario de consulta de Ubicaciones de PTL y BAC
' Converted from VB6 to VB.NET - Faithful line-by-line conversion
'
'Creación:      02/06/20
'
'Realización:   A.Esteban
'
'***********************************************************************

Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports ABG_Almacen_PTL.Modules
Imports ABG_Almacen_PTL.DataAccess

Public Class frmConsultaPTL
    Inherits Form

    ' ----- Constantes de Módulo (igual que VB6) -------------
    Private Const MOD_Nombre As String = "Consultas PTL"
    Private Const CML_Salir As Integer = 990
    Private Const LIS_ContenidoBAC As Integer = 1
    Private Const LIS_ContenidoCAJA As Integer = 2
    Private Const ColorRojo As Integer = &H8080FF
    Private Const ColorVerde As Integer = &H80FF80

    ' ----- Variables generales (igual que VB6) -------------
    Private ed As EntornoDeDatos
    Private dtArticulos As DataTable
    ' Private CustomDataFilter As clsDataFilter ' VB6 compatibility

    ' Controles principales
    Private FrameLectura As Panel
    Private WithEvents txtLecturaCodigo As TextBox
    Private cmdSalir As Button

    ' Frame datos
    Private fraArticulo As Panel

    ' Labels - títulos
    Private Label1 As Label
    Private Label2 As Label
    Private Label3 As Label
    Private Label4 As Label
    Private Label5 As Label
    Private Label6 As Label
    Private Label7 As Label
    Private Label8 As Label
    Private Label9 As Label

    ' Labels - valores
    Private lbUbicacion As Label
    Private lbBAC As Label
    Private lbGrupo As Label
    Private lbTablilla As Label
    Private lbNumCaja As Label
    Private lbUds As Label
    Private lbPeso As Label
    Private lbVolumen As Label
    Private lbTipoCaja As Label
    Private lbNombreCaja As Label

    ' Frame artículos
    Private fraArticulos As Panel
    Private dgvArticulos As DataGridView

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
        FrameLectura = New Panel()
        FrameLectura.BackColor = Color.Gray
        FrameLectura.Location = New Point(0, 0)
        FrameLectura.Size = New Size(254, 301)      ' 3810x4515 twips / 15

        ' txtLecturaCodigo
        txtLecturaCodigo = New TextBox()
        txtLecturaCodigo.Font = New Font("Arial", 14.25F)
        txtLecturaCodigo.BackColor = Color.White
        txtLecturaCodigo.Location = New Point(5, 5)     ' 75/15, 75/15
        txtLecturaCodigo.Size = New Size(180, 30)       ' 2700/15
        txtLecturaCodigo.MaxLength = 36

        ' cmdSalir (cmdAccion(990) en VB6)
        cmdSalir = New Button()
        cmdSalir.Text = "SALIR"
        cmdSalir.Location = New Point(190, 5)       ' 2850/15, 75/15
        cmdSalir.Size = New Size(61, 30)            ' 915/15, 450/15
        cmdSalir.TabStop = False
        AddHandler cmdSalir.Click, AddressOf cmdSalir_Click

        ' fraArticulo - Frame de datos
        fraArticulo = New Panel()
        fraArticulo.BackColor = Color.FromArgb(&HC0, &HC0, &HC0)
        fraArticulo.Location = New Point(5, 40)     ' 75/15, 600/15
        fraArticulo.Size = New Size(245, 130)       ' 3675/15, 1950/15

        ' Labels de etiquetas
        Label2 = New Label()
        Label2.Text = "Ubicación"
        Label2.Location = New Point(5, 5)
        Label2.Size = New Size(51, 19)

        Label3 = New Label()
        Label3.Text = "BAC"
        Label3.Location = New Point(5, 30)
        Label3.Size = New Size(51, 19)

        Label1 = New Label()
        Label1.Text = "Grupo"
        Label1.Location = New Point(5, 55)
        Label1.Size = New Size(51, 19)

        Label4 = New Label()
        Label4.Text = "Tablilla"
        Label4.Location = New Point(100, 55)
        Label4.Size = New Size(41, 19)

        Label8 = New Label()
        Label8.Text = "Caja"
        Label8.Location = New Point(175, 55)
        Label8.Size = New Size(31, 19)

        Label5 = New Label()
        Label5.Text = "Peso"
        Label5.Location = New Point(5, 80)
        Label5.Size = New Size(51, 19)

        Label6 = New Label()
        Label6.Text = "Volumen"
        Label6.Location = New Point(115, 80)
        Label6.Size = New Size(56, 19)

        Label7 = New Label()
        Label7.Text = "Tipo Caja"
        Label7.Location = New Point(5, 105)
        Label7.Size = New Size(51, 19)

        Label9 = New Label()
        Label9.Text = "Uds"
        Label9.Location = New Point(175, 105)
        Label9.Size = New Size(31, 19)

        ' Labels de valores
        lbUbicacion = New Label()
        lbUbicacion.BackColor = Color.White
        lbUbicacion.Font = New Font("MS Sans Serif", 9.75F, FontStyle.Bold)
        lbUbicacion.TextAlign = ContentAlignment.MiddleCenter
        lbUbicacion.Location = New Point(60, 5)
        lbUbicacion.Size = New Size(177, 19)

        lbBAC = New Label()
        lbBAC.BackColor = Color.White
        lbBAC.Font = New Font("MS Sans Serif", 9.75F, FontStyle.Bold)
        lbBAC.TextAlign = ContentAlignment.MiddleCenter
        lbBAC.Location = New Point(60, 30)
        lbBAC.Size = New Size(177, 19)

        lbGrupo = New Label()
        lbGrupo.BackColor = Color.White
        lbGrupo.TextAlign = ContentAlignment.MiddleCenter
        lbGrupo.Location = New Point(60, 55)
        lbGrupo.Size = New Size(37, 19)

        lbTablilla = New Label()
        lbTablilla.BackColor = Color.White
        lbTablilla.TextAlign = ContentAlignment.MiddleCenter
        lbTablilla.Location = New Point(145, 55)
        lbTablilla.Size = New Size(27, 19)

        lbNumCaja = New Label()
        lbNumCaja.BackColor = Color.White
        lbNumCaja.TextAlign = ContentAlignment.MiddleCenter
        lbNumCaja.Location = New Point(210, 55)
        lbNumCaja.Size = New Size(27, 19)

        lbPeso = New Label()
        lbPeso.BackColor = Color.White
        lbPeso.TextAlign = ContentAlignment.MiddleCenter
        lbPeso.Location = New Point(60, 80)
        lbPeso.Size = New Size(52, 19)

        lbVolumen = New Label()
        lbVolumen.BackColor = Color.White
        lbVolumen.TextAlign = ContentAlignment.MiddleCenter
        lbVolumen.Location = New Point(175, 80)
        lbVolumen.Size = New Size(62, 19)

        lbTipoCaja = New Label()
        lbTipoCaja.BackColor = Color.White
        lbTipoCaja.TextAlign = ContentAlignment.MiddleCenter
        lbTipoCaja.Location = New Point(60, 105)
        lbTipoCaja.Size = New Size(27, 19)

        lbNombreCaja = New Label()
        lbNombreCaja.BackColor = Color.White
        lbNombreCaja.TextAlign = ContentAlignment.MiddleCenter
        lbNombreCaja.Location = New Point(90, 105)
        lbNombreCaja.Size = New Size(81, 19)

        lbUds = New Label()
        lbUds.BackColor = Color.White
        lbUds.TextAlign = ContentAlignment.MiddleCenter
        lbUds.Location = New Point(210, 105)
        lbUds.Size = New Size(27, 19)

        ' Agregar labels al frame de datos
        fraArticulo.Controls.AddRange(New Control() {
            Label2, Label3, Label1, Label4, Label8, Label5, Label6, Label7, Label9,
            lbUbicacion, lbBAC, lbGrupo, lbTablilla, lbNumCaja, lbPeso, lbVolumen, lbTipoCaja, lbNombreCaja, lbUds
        })

        ' fraArticulos - Frame de lista de artículos
        fraArticulos = New Panel()
        fraArticulos.BackColor = Color.FromArgb(&HE0, &HE0, &HE0)
        fraArticulos.Location = New Point(5, 175)   ' 75/15, 2625/15
        fraArticulos.Size = New Size(245, 118)      ' 3675/15, 1770/15

        ' DataGridView para artículos (reemplaza UltraGrid)
        dgvArticulos = New DataGridView()
        dgvArticulos.Location = New Point(5, 5)
        dgvArticulos.Size = New Size(234, 109)
        dgvArticulos.AllowUserToAddRows = False
        dgvArticulos.AllowUserToDeleteRows = False
        dgvArticulos.ReadOnly = True
        dgvArticulos.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvArticulos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvArticulos.RowHeadersVisible = False
        dgvArticulos.BackgroundColor = Color.FromArgb(&HFF, &HDC, &HCE)
        dgvArticulos.Font = New Font("Arial", 8.25F)

        fraArticulos.Controls.Add(dgvArticulos)

        ' Agregar controles al FrameLectura
        FrameLectura.Controls.AddRange(New Control() {
            txtLecturaCodigo, cmdSalir, fraArticulo, fraArticulos
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

    Private Sub txtLecturaCodigo_GotFocus(sender As Object, e As EventArgs) Handles txtLecturaCodigo.GotFocus
        txtLecturaCodigo.BackColor = Color.FromArgb(&HC0, &HFF, &HC0)
    End Sub

    Private Sub txtLecturaCodigo_LostFocus(sender As Object, e As EventArgs) Handles txtLecturaCodigo.LostFocus
        txtLecturaCodigo.BackColor = SystemColors.Window
    End Sub

    Private Sub txtLecturaCodigo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtLecturaCodigo.KeyDown
        If e.KeyCode = Keys.Return Then
            ' Inicializa la visualización
            RefrescarDatos(True)

            Select Case txtLecturaCodigo.Text.Length
                Case 12 ' Unidad de transporte / Ubicación
                    ' Comprobar si la lectura es un BAC
                    Label3.Text = "BAC"
                    If fValidarBAC(txtLecturaCodigo.Text, False) = False Then
                        ' Comprobar si la lectura es una ubicación
                        If fValidarUbicacion(txtLecturaCodigo.Text, False) = False Then
                            ' No existe la ubicación / BAC
                            wsMensaje(" No se ha encontrado Ubicación o BAC", TipoMensaje.MENSAJE_Grave)
                        End If
                    End If

                Case 18 ' SSCC de Caja
                    Label3.Text = "CAJA"
                    fValidarCaja(txtLecturaCodigo.Text, True)

                Case 20 ' SSCC de Caja con prefijo
                    Label3.Text = "CAJA"
                    fValidarCaja(txtLecturaCodigo.Text.Substring(2), True)

            End Select

            txtLecturaCodigo.Text = ""
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Function fValidarBAC(ByVal stBAC As String, Optional ByVal blMensaje As Boolean = True) As Boolean
        Dim bCalculoPeso As Boolean
        Dim bCalculoVolumen As Boolean

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
                Dim unicod As String = If(Not IsDBNull(row("unicod")), row("unicod").ToString(), "")
                Dim uniest As Integer = If(Not IsDBNull(row("uniest")), CInt(row("uniest")), 0)
                Dim unigru As Integer = If(Not IsDBNull(row("unigru")), CInt(row("unigru")), 0)
                Dim unitab As Integer = If(Not IsDBNull(row("unitab")), CInt(row("unitab")), 0)
                Dim unicaj As String = If(Not IsDBNull(row("unicaj")), row("unicaj").ToString(), "")
                Dim tipdes As String = If(Not IsDBNull(row("tipdes")), row("tipdes").ToString(), "")
                Dim uninca As String = If(Not IsDBNull(row("uninca")), row("uninca").ToString(), "")

                bCalculoPeso = unipes > unipma
                bCalculoVolumen = univol > univma

                ' Se muestran los datos
                If IsDBNull(row("ubicod")) Then
                    RefrescarDatos(False, 0, 0, 0, 0, 0, unicod, uniest, unigru, unitab, unipes, univol, unicaj, tipdes, uninca, bCalculoPeso, bCalculoVolumen)
                Else
                    Dim ubicod As Integer = CInt(row("ubicod"))
                    Dim ubialm As Integer = If(Not IsDBNull(row("ubialm")), CInt(row("ubialm")), 0)
                    Dim ubiblo As Integer = If(Not IsDBNull(row("ubiblo")), CInt(row("ubiblo")), 0)
                    Dim ubifil As Integer = If(Not IsDBNull(row("ubifil")), CInt(row("ubifil")), 0)
                    Dim ubialt As Integer = If(Not IsDBNull(row("ubialt")), CInt(row("ubialt")), 0)
                    RefrescarDatos(False, ubicod, ubialm, ubiblo, ubifil, ubialt, unicod, uniest, unigru, unitab, unipes, univol, unicaj, tipdes, uninca, bCalculoPeso, bCalculoVolumen)
                End If

                ' Lista de artículos contenidos en el BAC
                sRefrescarArticulosBAC(unigru, unicod)
            Else
                If blMensaje Then wsMensaje(" No existe el BAC ", TipoMensaje.MENSAJE_Grave)
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

                ' Si existe comprueba si tiene un BAC asociado
                If IsDBNull(row("unicod")) Then
                    If blMensaje Then wsMensaje(" La Ubicación no tiene asociada un BAC ", TipoMensaje.MENSAJE_Grave)
                    lbUbicacion.Text = $"({ubicod}) {iALM:000}.{iBLO:000}.{iFIL:000}.{iALT:000}"
                Else
                    Dim unicod As String = row("unicod").ToString()
                    If fValidarBAC(unicod, False) = False Then
                        If blMensaje Then wsMensaje(" La Ubicación no tiene asociada un BAC válido ", TipoMensaje.MENSAJE_Grave)
                    End If
                End If
            Else
                If blMensaje Then wsMensaje(" No existe la Unidad de Transporte ", TipoMensaje.MENSAJE_Grave)
                lbUbicacion.Text = ""
            End If
        Catch ex As Exception
            If blMensaje Then wsMensaje($" Error: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        Return fValidarUbicacion
    End Function

    Private Function fValidarCaja(ByVal stSSCC As String, Optional ByVal blMensaje As Boolean = True) As Boolean
        Dim bCalculoPeso As Boolean
        Dim bCalculoVolumen As Boolean
        Dim bEstado As Integer

        fValidarCaja = False

        Try
            Dim dtCaja As DataTable = ed.DameDatosCAJAdePTL(stSSCC)

            If dtCaja.Rows.Count > 0 Then
                fValidarCaja = True
                Dim row As DataRow = dtCaja.Rows(0)

                Dim ltcvol As Double = If(Not IsDBNull(row("ltcvol")), CDbl(row("ltcvol")), 0)
                Dim ltcssc As String = If(Not IsDBNull(row("ltcssc")), row("ltcssc").ToString(), "")
                Dim ltcgru As Integer = If(Not IsDBNull(row("ltcgru")), CInt(row("ltcgru")), 0)
                Dim ltctab As Integer = If(Not IsDBNull(row("ltctab")), CInt(row("ltctab")), 0)
                Dim ltcpes As Double = If(Not IsDBNull(row("ltcpes")), CDbl(row("ltcpes")), 0)
                Dim ltctip As String = If(Not IsDBNull(row("ltctip")), row("ltctip").ToString(), "")
                Dim tipdes As String = If(Not IsDBNull(row("tipdes")), row("tipdes").ToString(), "")
                Dim ltccaj As String = If(Not IsDBNull(row("ltccaj")), row("ltccaj").ToString(), "")

                bEstado = If(ltcvol > 0, 1, 0)

                ' Se muestran los datos
                RefrescarDatos(False, 0, 0, 0, 0, 0, ltcssc, bEstado, ltcgru, ltctab, ltcpes, ltcvol, ltctip, tipdes, ltccaj, bCalculoPeso, bCalculoVolumen)

                ' Lista de artículos contenidos en la CAJA
                sRefrescarArticulosCAJA(ltcgru, ltctab, ltccaj)
            Else
                If blMensaje Then wsMensaje(" No existe la CAJA ", TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            If blMensaje Then wsMensaje($" Error: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        Return fValidarCaja
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
                               Optional sNumCaja As String = "",
                               Optional bPeso As Boolean = False,
                               Optional bVolumen As Boolean = False)

        If sEnBlanco = True Then
            ' Inicia la visualización
            lbUbicacion.Text = ""
            lbBAC.Text = ""
            lbBAC.BackColor = Color.White

            lbGrupo.Text = ""
            lbTablilla.Text = ""
            lbUds.Text = ""

            lbPeso.Text = ""
            lbPeso.BackColor = Color.White

            lbVolumen.Text = ""
            lbVolumen.BackColor = Color.White

            lbTipoCaja.Text = ""
            lbNombreCaja.Text = ""
            lbNumCaja.Text = ""

            dgvArticulos.DataSource = Nothing
        Else
            If sCodUbicacion = 0 Then
                lbUbicacion.Text = "SIN UBICACION"
                lbUbicacion.Text = "-------------"
            Else
                lbUbicacion.Text = $"({sCodUbicacion}) {sALM:000}.{sBLO:000}.{sFIL:000}.{sALT:000}"
            End If

            lbBAC.Text = sBAC
            lbBAC.BackColor = If(sEstadoBAC = 0, Color.White, Color.FromArgb((ColorVerde >> 16) And &HFF, (ColorVerde >> 8) And &HFF, ColorVerde And &HFF))

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
            lbNumCaja.Text = sNumCaja
        End If
    End Sub

    Private Sub sRefrescarArticulosBAC(ByVal sGrupo As Long, sBAC As String)
        Dim iUds As Integer = 0

        dgvArticulos.DataSource = Nothing

        Try
            Dim dtContenido As DataTable = ed.DameContenidoBacGrupo(CInt(sGrupo), sBAC)

            If dtContenido.Rows.Count > 0 Then
                ' Calcular unidades
                For Each row As DataRow In dtContenido.Rows
                    If Not IsDBNull(row("unican")) Then
                        iUds = iUds + CInt(row("unican"))
                    End If
                Next

                ' Configurar columnas del DataGridView
                ConfigurarColumnasBAC()
                dgvArticulos.DataSource = dtContenido
            End If
        Catch ex As Exception
            wsMensaje($" Error al cargar artículos: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        lbUds.Text = iUds.ToString()
    End Sub

    Private Sub sRefrescarArticulosCAJA(ByVal sGrupo As Long, sTablilla As Long, sCaja As String)
        Dim iUds As Integer = 0

        dgvArticulos.DataSource = Nothing

        Try
            Dim dtContenido As DataTable = ed.DameContenidoCajaGrupo(CInt(sGrupo), CInt(sTablilla), sCaja)

            If dtContenido.Rows.Count > 0 Then
                ' Calcular unidades
                For Each row As DataRow In dtContenido.Rows
                    If Not IsDBNull(row("ltacan")) Then
                        iUds = iUds + CInt(CDbl(row("ltacan")))
                    End If
                Next

                ' Configurar columnas del DataGridView
                ConfigurarColumnasCAJA()
                dgvArticulos.DataSource = dtContenido
            End If
        Catch ex As Exception
            wsMensaje($" Error al cargar artículos: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        lbUds.Text = iUds.ToString()
    End Sub

    Private Sub ConfigurarColumnasBAC()
        ' Ocultar todas las columnas y mostrar solo las necesarias
        dgvArticulos.AutoGenerateColumns = False

        dgvArticulos.Columns.Clear()

        Dim colCodigo As New DataGridViewTextBoxColumn()
        colCodigo.DataPropertyName = "uniart"
        colCodigo.HeaderText = "Codigo"
        colCodigo.Width = 60
        dgvArticulos.Columns.Add(colCodigo)

        Dim colNombre As New DataGridViewTextBoxColumn()
        colNombre.DataPropertyName = "artnom"
        colNombre.HeaderText = "Articulo"
        colNombre.Width = 120
        dgvArticulos.Columns.Add(colNombre)

        Dim colCant As New DataGridViewTextBoxColumn()
        colCant.DataPropertyName = "unican"
        colCant.HeaderText = "Cant"
        colCant.Width = 50
        dgvArticulos.Columns.Add(colCant)
    End Sub

    Private Sub ConfigurarColumnasCAJA()
        dgvArticulos.AutoGenerateColumns = False

        dgvArticulos.Columns.Clear()

        Dim colCodigo As New DataGridViewTextBoxColumn()
        colCodigo.DataPropertyName = "ltaart"
        colCodigo.HeaderText = "Codigo"
        colCodigo.Width = 60
        dgvArticulos.Columns.Add(colCodigo)

        Dim colNombre As New DataGridViewTextBoxColumn()
        colNombre.DataPropertyName = "artnom"
        colNombre.HeaderText = "Articulo"
        colNombre.Width = 120
        dgvArticulos.Columns.Add(colNombre)

        Dim colCant As New DataGridViewTextBoxColumn()
        colCant.DataPropertyName = "ltacan"
        colCant.HeaderText = "Cant"
        colCant.Width = 50
        dgvArticulos.Columns.Add(colCant)
    End Sub

End Class
