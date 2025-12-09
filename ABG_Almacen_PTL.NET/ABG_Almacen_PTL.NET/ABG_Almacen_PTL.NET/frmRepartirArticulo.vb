'***********************************************************************
'Nombre: frmRepartirArticulo.vb
' Formulario para el reparto de Artículos en BACs casilleros de PTL
' Converted from VB6 to VB.NET - Faithful line-by-line conversion
'
'Creación:      02/09/20
'
'Realización:   A.Esteban
'
'***********************************************************************

Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports ABG_Almacen_PTL.Modules
Imports ABG_Almacen_PTL.DataAccess
Imports ABG_Almacen_PTL.Classes

' ----- Constantes de Módulo -------------
Public Class frmRepartirArticulo
    Inherits Form

    Private Const MOD_Nombre As String = "Repartir Articulo"
    Private Const CML_Salir As Integer = 990
    Private Const LIS_ContenidoBAC As Integer = 1
    Private Const ColorRojo As Integer = &H8080FF
    Private Const ColorVerde As Integer = &H80FF80

    ' ----- Variables generales (igual que VB6) -------------
    Private ed As EntornoDeDatos
    Private tUsuario As Integer
    Private bInicio As Boolean
    ' Private CustomDataFilter As clsDataFilter ' VB6 compatibility (not implemented)

    ' Controles principales
    Private FrameLectura As Panel
    Private WithEvents txtLecturaCodigo As TextBox
    Private cmdSalir As Button  ' Botón SALIR (VB6: cmdAccion(990))
    Private WithEvents Combo1 As ComboBox
    Private pColor As Panel

    ' Labels
    Private lbTexto As Label
    Private Label8 As Label
    Private Label4 As Label
    Private Label2 As Label
    Private Label14 As Label
    Private Label15 As Label
    Private Label1 As Label
    Private Label3 As Label

    Private lbArticulo As Label
    Private lbNombreArticulo As Label
    Private lbEAN13 As Label
    Private lbSTD As Label
    Private lbPeso As Label
    Private lbVolumen As Label

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
        Me.Text = "Form1"                           ' VB6: Caption = "Form1"
        Me.ClientSize = New Size(273, 301)         ' VB6: 4095x4515 twips / 15
        Me.KeyPreview = True                        ' VB6: KeyPreview = -1 'True
        Me.ShowInTaskbar = False                    ' VB6: ShowInTaskbar = 0 'False
        Me.StartPosition = FormStartPosition.Manual

        ' FrameLectura (PictureBox en VB6)
        ' VB6: BackColor = &H00808080& = Gray
        FrameLectura = New Panel()
        FrameLectura.BackColor = Color.Gray
        FrameLectura.Location = New Point(0, 0)
        FrameLectura.Size = New Size(254, 301)      ' 3810x4515 twips / 15

        ' lbTexto - Título
        lbTexto = New Label()
        lbTexto.Text = "REPARTIR ARTICULO"
        lbTexto.Font = New Font("MS Sans Serif", 12, FontStyle.Bold)
        lbTexto.ForeColor = Color.White
        lbTexto.BackColor = Color.Gray
        lbTexto.TextAlign = ContentAlignment.MiddleCenter
        lbTexto.Location = New Point(5, 5)         ' 75/15, 75/15
        lbTexto.Size = New Size(245, 16)           ' 3675/15, 240/15

        ' txtLecturaCodigo
        txtLecturaCodigo = New TextBox()
        txtLecturaCodigo.Font = New Font("Arial", 14.25F)
        txtLecturaCodigo.BackColor = Color.White
        txtLecturaCodigo.Location = New Point(10, 30)   ' 150/15, 450/15
        txtLecturaCodigo.Size = New Size(236, 30)       ' 3540/15
        txtLecturaCodigo.MaxLength = 36

        ' Label8 - Instrucción
        Label8 = New Label()
        Label8.Text = "Leer o teclear Artículo"
        Label8.BackColor = Color.Gray
        Label8.ForeColor = Color.White
        Label8.TextAlign = ContentAlignment.MiddleCenter
        Label8.Location = New Point(10, 60)        ' 150/15, 900/15
        Label8.Size = New Size(236, 16)            ' 3540/15, 240/15

        ' Label4 - Puesto
        Label4 = New Label()
        Label4.Text = "Puesto"
        Label4.Location = New Point(10, 80)        ' 150/15, 1200/15
        Label4.Size = New Size(51, 21)             ' 765/15, 315/15

        ' Combo1 - Puestos
        Combo1 = New ComboBox()
        Combo1.DropDownStyle = ComboBoxStyle.DropDownList
        Combo1.Location = New Point(65, 80)        ' 975/15
        Combo1.Size = New Size(141, 21)            ' 2115/15

        ' pColor
        pColor = New Panel()
        pColor.BackColor = SystemColors.Window
        pColor.Location = New Point(210, 80)       ' 3150/15
        pColor.Size = New Size(36, 21)             ' 540/15

        ' Label2 - Código
        Label2 = New Label()
        Label2.Text = "Código"
        Label2.Location = New Point(10, 105)       ' 150/15, 1575/15
        Label2.Size = New Size(51, 21)

        ' lbArticulo
        lbArticulo = New Label()
        lbArticulo.BackColor = Color.White
        lbArticulo.TextAlign = ContentAlignment.MiddleCenter
        lbArticulo.Location = New Point(63, 105)   ' 945/15
        lbArticulo.Size = New Size(183, 21)        ' 2745/15

        ' lbNombreArticulo
        lbNombreArticulo = New Label()
        lbNombreArticulo.BackColor = Color.White
        lbNombreArticulo.TextAlign = ContentAlignment.MiddleCenter
        lbNombreArticulo.Location = New Point(10, 130)   ' 150/15, 1950/15
        lbNombreArticulo.Size = New Size(236, 47)        ' 3540/15, 705/15

        ' Label14 - EAN13
        Label14 = New Label()
        Label14.Text = "EAN13"
        Label14.Location = New Point(10, 180)      ' 150/15, 2700/15
        Label14.Size = New Size(51, 21)

        ' lbEAN13
        lbEAN13 = New Label()
        lbEAN13.BackColor = Color.White
        lbEAN13.TextAlign = ContentAlignment.MiddleCenter
        lbEAN13.Location = New Point(63, 180)
        lbEAN13.Size = New Size(183, 21)

        ' Label15 - STD
        Label15 = New Label()
        Label15.Text = "STD"
        Label15.Location = New Point(10, 205)      ' 150/15, 3075/15
        Label15.Size = New Size(51, 21)

        ' lbSTD
        lbSTD = New Label()
        lbSTD.BackColor = Color.White
        lbSTD.TextAlign = ContentAlignment.MiddleCenter
        lbSTD.Location = New Point(63, 205)
        lbSTD.Size = New Size(47, 21)              ' 705/15

        ' Label1 - Peso
        Label1 = New Label()
        Label1.Text = "Peso"
        Label1.Location = New Point(10, 230)       ' 150/15, 3450/15
        Label1.Size = New Size(51, 21)

        ' lbPeso
        lbPeso = New Label()
        lbPeso.BackColor = Color.White
        lbPeso.TextAlign = ContentAlignment.MiddleCenter
        lbPeso.Location = New Point(63, 230)
        lbPeso.Size = New Size(47, 21)

        ' Label3 - Volumen
        Label3 = New Label()
        Label3.Text = "Volumen"
        Label3.Location = New Point(125, 230)      ' 1875/15
        Label3.Size = New Size(51, 21)

        ' lbVolumen
        lbVolumen = New Label()
        lbVolumen.BackColor = Color.White
        lbVolumen.TextAlign = ContentAlignment.MiddleCenter
        lbVolumen.Location = New Point(180, 230)   ' 2700/15
        lbVolumen.Size = New Size(47, 21)

        ' cmdSalir - SALIR (VB6: cmdAccion(990))
        cmdSalir = New Button()
        cmdSalir.Text = "SALIR"
        cmdSalir.Location = New Point(5, 260)    ' 75/15, 3900/15
        cmdSalir.Size = New Size(244, 35)        ' 3660/15, 525/15
        cmdSalir.TabStop = False
        AddHandler cmdSalir.Click, AddressOf cmdSalir_Click

        ' Agregar controles al FrameLectura
        FrameLectura.Controls.AddRange(New Control() {
            lbTexto, txtLecturaCodigo, Label8, Label4, Combo1, pColor,
            Label2, lbArticulo, lbNombreArticulo,
            Label14, lbEAN13, Label15, lbSTD,
            Label1, lbPeso, Label3, lbVolumen,
            cmdSalir
        })

        Me.Controls.Add(FrameLectura)

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    '*******************************************************************************

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bInicio = True
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

        CargarPistolas()
        tUsuario = 0

        FrameLectura.Left = 0

        Cursor = Cursors.Default
        bInicio = False
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

    Private Sub CargarPistolas()
        ' Combo de puestos de trabajo
        Combo1.Items.Clear()
        Combo1.Items.Add("(0) Sin puesto")

        Try
            Dim dtPuestos As DataTable = ed.DamePuestosTrabajoPTL()
            If dtPuestos IsNot Nothing Then
                For Each row As DataRow In dtPuestos.Rows
                    Dim puecod As Integer = If(Not IsDBNull(row("puecod")), CInt(row("puecod")), 0)
                    Dim puedes As String = If(Not IsDBNull(row("puedes")), row("puedes").ToString(), "")
                    Dim puecol As Integer = If(Not IsDBNull(row("puecol")), CInt(row("puecol")), 0)
                    Combo1.Items.Add($"({puecod}) {puedes}")
                Next
            End If
        Catch
            ' Si falla la carga, usar puestos por defecto
            For i As Integer = 1 To 7
                Combo1.Items.Add($"({i}) Puesto {i}")
            Next
        End Try

        If Combo1.Items.Count > 0 Then
            Combo1.SelectedIndex = 0
        End If
    End Sub

    Private Sub Combo1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo1.SelectedIndexChanged
        ' Aplica el color y el usuario
        Select Case Combo1.SelectedIndex
            Case 1
                pColor.BackColor = Color.White
                tUsuario = ExtractUserFromCombo()
            Case 2
                pColor.BackColor = Color.Yellow
                tUsuario = ExtractUserFromCombo()
            Case 3
                pColor.BackColor = Color.Magenta
                tUsuario = ExtractUserFromCombo()
            Case 4
                pColor.BackColor = Color.Cyan
                tUsuario = ExtractUserFromCombo()
            Case 5
                pColor.BackColor = Color.Blue
                tUsuario = ExtractUserFromCombo()
            Case 6
                pColor.BackColor = Color.Green
                tUsuario = ExtractUserFromCombo()
            Case 7
                pColor.BackColor = Color.Red
                tUsuario = ExtractUserFromCombo()
            Case Else
                pColor.BackColor = Color.FromArgb(&H80, &H80, &H80)
                tUsuario = 0
        End Select

        If Not bInicio Then txtLecturaCodigo.Focus()
    End Sub

    Private Function ExtractUserFromCombo() As Integer
        Try
            Dim texto As String = Combo1.Text
            If texto.StartsWith("(") Then
                Dim endPos As Integer = texto.IndexOf(")"c)
                If endPos > 1 Then
                    Return CInt(texto.Substring(1, endPos - 1))
                End If
            End If
        Catch
        End Try
        Return 0
    End Function

    Private Sub txtLecturaCodigo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtLecturaCodigo.KeyDown
        If e.KeyCode = Keys.Return Then
            ' Inicializa la visualización
            RefrescarDatos(True)

            Select Case txtLecturaCodigo.Text.Length
                Case 13 ' EAN13
                    ' Comprobar si la lectura es un EAN13
                    If fValidarEAN13(txtLecturaCodigo.Text, True) = False Then
                        wsMensaje(" No se ha encontrado el Artículo", TipoMensaje.MENSAJE_Grave)
                    End If

                Case 4, 5 ' Código de artículo
                    ' Comprobar si la lectura es un Artículo
                    If fValidarArticulo(txtLecturaCodigo.Text, True) = False Then
                        wsMensaje(" No se ha encontrado el Artículo", TipoMensaje.MENSAJE_Grave)
                    End If
            End Select

            txtLecturaCodigo.Text = ""
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Function fValidarArticulo(ByVal stArticulo As String, Optional ByVal blMensaje As Boolean = True) As Boolean
        fValidarArticulo = False

        Try
            Dim dtArticulo As DataTable = ed.DameArticuloConsulta(stArticulo)

            If dtArticulo.Rows.Count > 0 Then
                fValidarArticulo = True
                Dim row As DataRow = dtArticulo.Rows(0)

                ' Se muestran los datos
                Dim artcod As Long = If(Not IsDBNull(row("artcod")), CLng(row("artcod")), 0)
                Dim artnom As String = If(Not IsDBNull(row("artnom")), row("artnom").ToString(), "")
                Dim artean As String = If(Not IsDBNull(row("artean")), row("artean").ToString(), "")
                Dim artcj3 As Integer = If(Not IsDBNull(row("artcj3")), CInt(row("artcj3")), 0)
                Dim artpea As Double = If(Not IsDBNull(row("artpea")), CDbl(row("artpea")), 0)
                Dim artcua As Double = If(Not IsDBNull(row("artcua")), CDbl(row("artcua")), 0)

                RefrescarDatos(False, artcod, artnom, artean, artcj3, artpea, artcua)

                ' Se procede a repartir el artículo
                If RepartirArticulo(artcod) Then
                    wsMensaje($" Se ha reservado el BAC para el Artículo: {artcod}", TipoMensaje.MENSAJE_Exclamacion)
                End If
            Else
                If blMensaje Then wsMensaje(" No existe el Artículo ", TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            If blMensaje Then wsMensaje($" Error: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        Return fValidarArticulo
    End Function

    Private Function fValidarEAN13(ByVal stEAN13 As String, Optional ByVal blMensaje As Boolean = True) As Boolean
        fValidarEAN13 = False
        Dim tEAN13 As String = stEAN13.Substring(0, Math.Min(13, stEAN13.Length))

        Try
            Dim dtArticulo As DataTable = ed.DameArticuloEAN13(stEAN13)

            If dtArticulo.Rows.Count > 0 Then
                fValidarEAN13 = True
                Dim row As DataRow = dtArticulo.Rows(0)

                Dim artcod As Long = If(Not IsDBNull(row("artcod")), CLng(row("artcod")), 0)
                Dim artnom As String = If(Not IsDBNull(row("artnom")), row("artnom").ToString(), "")
                Dim artean As String = If(Not IsDBNull(row("artean")), row("artean").ToString(), "")
                Dim artcj3 As Integer = If(Not IsDBNull(row("artcj3")), CInt(row("artcj3")), 0)
                Dim artpea As Double = If(Not IsDBNull(row("artpea")), CDbl(row("artpea")), 0)
                Dim artcua As Double = If(Not IsDBNull(row("artcua")), CDbl(row("artcua")), 0)

                RefrescarDatos(False, artcod, artnom, artean, artcj3, artpea, artcua)

                If RepartirArticulo(artcod) Then
                    wsMensaje($" Se ha reservado el BAC para el Artículo: {artcod}", TipoMensaje.MENSAJE_Exclamacion)
                End If
            Else
                If blMensaje Then wsMensaje(" No existe el Artículo ", TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            If blMensaje Then wsMensaje($" Error: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        Return fValidarEAN13
    End Function

    Private Sub RefrescarDatos(sEnBlanco As Boolean,
                               Optional sArticulo As Long = 0,
                               Optional sNombre As String = "",
                               Optional sEAN13 As String = "",
                               Optional sSTD As Integer = 0,
                               Optional sPeso As Double = 0,
                               Optional sVolumen As Double = 0)

        If sEnBlanco = True Then
            lbArticulo.Text = ""
            lbNombreArticulo.Text = ""
            lbEAN13.Text = ""
            lbSTD.Text = ""
            lbPeso.Text = ""
            lbVolumen.Text = ""
        Else
            lbArticulo.Text = sArticulo.ToString()
            lbNombreArticulo.Text = sNombre
            lbEAN13.Text = sEAN13
            lbSTD.Text = sSTD.ToString()
            lbPeso.Text = sPeso.ToString("#0.0000")
            lbVolumen.Text = sVolumen.ToString("#0.0000")
        End If
    End Sub

    Private Function RepartirArticulo(tArticulo As Long) As Boolean
        ' Reparto del Artículo
        Dim Retorno As Integer = 0
        Dim msgSalida As String = ""

        RepartirArticulo = False

        Try
            ed.ReservaBACdePTL(tArticulo, tUsuario, Retorno, msgSalida)

            If Retorno = 0 Then
                RepartirArticulo = True
            Else
                wsMensaje($" No se ha podido repartir el Artículo. {msgSalida}", TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            wsMensaje($" Error al repartir: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        Return RepartirArticulo
    End Function

End Class
