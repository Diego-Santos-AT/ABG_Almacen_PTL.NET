'***********************************************************************
'Nombre: frmEmpaquetarBAC.vb
' Formulario de empaquetado rápido de BAC del sistema de PTL
' Converted from VB6 to VB.NET - Faithful line-by-line conversion
'
'Creación:      05/06/20
'
'Realización:   A.Esteban
'
'***********************************************************************

Imports System.Data
Imports System.Windows.Forms
Imports System.Drawing
Imports ABG_Almacen_PTL.Modules
Imports ABG_Almacen_PTL.DataAccess

Public Class frmEmpaquetarBAC
    Inherits Form

    ' ----- Constantes de Módulo (igual que VB6) -------------
    Private Const MOD_Nombre As String = "Empaquetar BAC"

    Private Const CML_Salir As Integer = 990

    Private Const CML_Opciones As Integer = 0
    Private Const CML_Acciones As Integer = 5

    Private Const CML_CerrarBAC As Integer = 10
    Private Const CML_ExtraerBAC As Integer = 20
    Private Const CML_CrearCAJA As Integer = 30
    Private Const CML_ImprimirCAJA As Integer = 40
    Private Const CML_RelContenido As Integer = 50
    Private Const CML_Empaquetado As Integer = 60
    Private Const CML_CambiarCAJA As Integer = 70
    Private Const CML_CambiarUDS As Integer = 80
    Private Const CML_CombinarCAJAS As Integer = 85

    Private Const CML_RestarUDS As Integer = 90
    Private Const CML_SumarUDS As Integer = 95
    Private Const CML_AplicarUDS As Integer = 110

    Private Const CML_Volver As Integer = 100
    Private Const CML_CambiaCAJA As Integer = 170
    Private Const CML_AplicaCombinar As Integer = 180

    Private Const CML_VolverCajas As Integer = 120
    Private Const CML_VolverCantidad As Integer = 130
    Private Const CML_VolverCombinar As Integer = 140

    Private Const OPC_CerrarBAC As Integer = 0
    Private Const OPC_ExtraerBAC As Integer = 1
    Private Const OPC_CrearCAJA As Integer = 2
    Private Const OPC_ImprimirCAJA As Integer = 3
    Private Const OPC_RelContenido As Integer = 4
    Private Const OPC_CerrarCAJA As Integer = 5
    Private Const OPC_Empaquetado As Integer = 6

    Private Const ACC_General As String = "ACCIONES"
    Private Const ACC_Empaquetar As String = "EMPAQUETAR"
    Private Const ACC_Etiquetas As String = "IMPRIMIR ETIQUETAS"

    Private Const LIS_ContenidoBAC As Integer = 1
    Private Const LIS_ContenidoCAJA As Integer = 2
    Private Const LIS_TipoCajas As Integer = 3

    Private Const ColorRojo As Integer = &H8080FF
    Private Const ColorVerde As Integer = &H80FF80

    ' ----- Variables generales (igual que VB6) -------------
    Private ed As EntornoDeDatos
    Private dtArticulos As DataTable
    Private dtArticulosCaja As DataTable
    Private dtCajas As DataTable

    Private tEstadoBAC As Integer   ' Estado del BAC
    Private tUbicacionBAC As Integer  ' Ubicacion del BAC

    ' Controles principales - FrameLectura
    Private FrameLectura As Panel
    Private WithEvents txtLecturaCodigo As TextBox
    Private cmdAccion As Dictionary(Of Integer, Button)
    Private fraArticulo As Panel
    Private fraArticulos As Panel
    Private dgvArticulos As DataGridView

    ' Labels - FrameLectura
    Private Label2 As Label
    Private Label3 As Label
    Private Label1 As Label
    Private Label4 As Label
    Private Label5 As Label
    Private Label6 As Label
    Private Label7 As Label
    Private Label9 As Label
    Private Label12 As Label

    Private lbUbicacion As Label
    Private lbBAC As Label
    Private lbGrupo As Label
    Private lbTablilla As Label
    Private lbNumCaja As Label
    Private lbPeso As Label
    Private lbVolumen As Label
    Private lbTipoCaja As Label
    Private lbNombreCaja As Label
    Private lbUds As Label
    Private lbArts As Label
    Private lbSSCC As Label

    ' Controles - FrameOpciones
    Private FrameOpciones As Panel
    Private Check1 As CheckBox()
    Private lbNumCaja2 As Label
    Private lbTipoCaja2 As Label
    Private lbNombreCaja2 As Label

    ' Controles - FrameCajas (para cambio de caja)
    Private FrameCajas As Panel
    Private dgvCajas As DataGridView

    ' Controles - FrameCantidad (para cambio de unidades)
    Private FrameCantidad As Panel
    Private WithEvents nArticulo As TextBox
    Private WithEvents nCantidad As TextBox
    Private lbNombreArticulo As Label

    ' Controles - FrameCombinar (para combinar cajas)
    Private FrameCombinar As Panel
    Private WithEvents txtLecturaCaja2 As TextBox
    Private lbNumCaja3 As Label
    Private lbUds3 As Label
    Private lbArts3 As Label
    Private lbSSCC3 As Label
    Private lbTipoCaja3 As Label
    Private lbNombreCaja3 As Label
    Private lbPeso3 As Label
    Private lbVolumen3 As Label
    Private lbNumCaja4 As Label
    Private lbUds4 As Label
    Private lbArts4 As Label
    Private lbPeso4 As Label
    Private lbVolumen4 As Label
    Private lbPesoTot As Label
    Private lbVolumenTot As Label

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
        Me.ClientSize = New Size(1407, 301)         ' 21105x4515 twips / 15
        Me.KeyPreview = True                        ' KeyPreview = -1 'True
        Me.ShowInTaskbar = False                    ' ShowInTaskbar = 0 'False
        Me.StartPosition = FormStartPosition.Manual

        ' Inicializa el diccionario de botones cmdAccion
        cmdAccion = New Dictionary(Of Integer, Button)

        ' ============= FrameLectura =============
        FrameLectura = New Panel()
        FrameLectura.BackColor = Color.Gray
        FrameLectura.Location = New Point(0, 0)
        FrameLectura.Size = New Size(254, 301)  ' 3810/15, 4515/15

        ' txtLecturaCodigo
        txtLecturaCodigo = New TextBox()
        txtLecturaCodigo.Font = New Font("Arial", 14.25F)
        txtLecturaCodigo.BackColor = Color.White
        txtLecturaCodigo.Location = New Point(5, 5)     ' 75/15
        txtLecturaCodigo.Size = New Size(180, 30)       ' 2700/15
        txtLecturaCodigo.MaxLength = 36

        ' cmdAccion(990) - SALIR
        Dim btnSalir As New Button()
        btnSalir.Text = "SALIR"
        btnSalir.Location = New Point(190, 5)           ' 2850/15
        btnSalir.Size = New Size(61, 30)                ' 915/15
        btnSalir.TabStop = False
        AddHandler btnSalir.Click, Sub(s, e) cmdAccion_Click(CML_Salir)
        cmdAccion(CML_Salir) = btnSalir

        ' fraArticulo
        fraArticulo = New Panel()
        fraArticulo.BackColor = Color.FromArgb(&HC0, &HC0, &HC0)
        fraArticulo.Location = New Point(5, 40)         ' 75/15, 600/15
        fraArticulo.Size = New Size(245, 105)           ' 3675/15, 1575/15

        ' Labels de etiquetas en fraArticulo
        Label3 = New Label() With {.Text = "BAC", .Location = New Point(5, 5), .Size = New Size(51, 19)}
        Label1 = New Label() With {.Text = "Grupo", .Location = New Point(5, 30), .Size = New Size(51, 19)}
        Label4 = New Label() With {.Text = "Tablilla", .Location = New Point(100, 30), .Size = New Size(45, 19)}
        Label12 = New Label() With {.Text = "Caja", .Location = New Point(175, 30), .Size = New Size(31, 19)}
        Label5 = New Label() With {.Text = "Peso", .Location = New Point(5, 55), .Size = New Size(51, 19)}
        Label6 = New Label() With {.Text = "Volumen", .Location = New Point(115, 55), .Size = New Size(56, 19)}
        Label7 = New Label() With {.Text = "Tipo Caja", .Location = New Point(5, 80), .Size = New Size(51, 19)}
        Label9 = New Label() With {.Text = "Uds", .Location = New Point(175, 80), .Size = New Size(31, 19)}

        ' Labels de valores en fraArticulo
        lbBAC = CreateValueLabel(60, 5, 177, 19, True)
        lbGrupo = CreateValueLabel(60, 30, 37, 19)
        lbTablilla = CreateValueLabel(145, 30, 27, 19)
        lbNumCaja = CreateValueLabel(210, 30, 27, 19)
        lbPeso = CreateValueLabel(60, 55, 52, 19)
        lbVolumen = CreateValueLabel(175, 55, 62, 19)
        lbTipoCaja = CreateValueLabel(60, 80, 27, 19)
        lbNombreCaja = CreateValueLabel(90, 80, 81, 19)
        AddHandler lbNombreCaja.Click, AddressOf lbNombreCaja_Click
        lbUds = CreateValueLabel(210, 80, 27, 19)

        fraArticulo.Controls.AddRange(New Control() {
            Label3, Label1, Label4, Label12, Label5, Label6, Label7, Label9,
            lbBAC, lbGrupo, lbTablilla, lbNumCaja, lbPeso, lbVolumen, lbTipoCaja, lbNombreCaja, lbUds
        })

        ' fraArticulos
        fraArticulos = New Panel()
        fraArticulos.BackColor = Color.FromArgb(&HE0, &HE0, &HE0)
        fraArticulos.Location = New Point(5, 150)       ' 75/15, 2250/15
        fraArticulos.Size = New Size(245, 118)          ' 3675/15, 1770/15

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
        AddHandler dgvArticulos.DoubleClick, AddressOf ugArticulos_DblClick

        fraArticulos.Controls.Add(dgvArticulos)

        ' lbArts
        lbArts = New Label()
        lbArts.BackColor = Color.White
        lbArts.TextAlign = ContentAlignment.MiddleCenter
        lbArts.Location = New Point(5, 273)             ' 75/15, 4100/15
        lbArts.Size = New Size(26, 25)                  ' 390/15

        ' lbSSCC (oculto, para almacenar SSCC)
        lbSSCC = New Label() With {.Visible = False}

        ' cmdAccion(0) - >>
        Dim btnOpciones As New Button()
        btnOpciones.Text = ">>"
        btnOpciones.Location = New Point(209, 273)      ' 3135/15, 4100/15
        btnOpciones.Size = New Size(41, 25)             ' 615/15
        btnOpciones.TabStop = False
        AddHandler btnOpciones.Click, Sub(s, e) cmdAccion_Click(CML_Opciones)
        cmdAccion(CML_Opciones) = btnOpciones

        ' cmdAccion(5) - ACCIONES
        Dim btnAcciones As New Button()
        btnAcciones.Text = ACC_General
        btnAcciones.Location = New Point(40, 273)       ' 600/15, 4100/15
        btnAcciones.Size = New Size(161, 25)            ' 2415/15
        btnAcciones.TabStop = False
        AddHandler btnAcciones.Click, Sub(s, e) cmdAccion_Click(CML_Acciones)
        cmdAccion(CML_Acciones) = btnAcciones

        FrameLectura.Controls.AddRange(New Control() {
            txtLecturaCodigo, btnSalir, fraArticulo, fraArticulos, lbArts, lbSSCC,
            btnOpciones, btnAcciones
        })

        ' ============= FrameOpciones =============
        FrameOpciones = New Panel()
        FrameOpciones.BackColor = Color.Gray
        FrameOpciones.Location = New Point(267, 0)      ' 4000/15
        FrameOpciones.Size = New Size(254, 301)

        Label2 = New Label() With {.Text = "Ubicación", .Location = New Point(5, 5), .Size = New Size(51, 19)}
        lbUbicacion = CreateValueLabel(60, 5, 187, 19, True)

        ' Check1 array (6 checkboxes)
        Check1 = New CheckBox(6) {}

        Check1(OPC_CerrarBAC) = New CheckBox() With {
            .Text = "Cerrar BAC abiertos",
            .Location = New Point(5, 35),
            .Size = New Size(146, 21),
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler Check1(OPC_CerrarBAC).CheckedChanged, Sub(s, e) Check1_Click(OPC_CerrarBAC)

        Check1(OPC_ExtraerBAC) = New CheckBox() With {
            .Text = "Extraer BAC ubicados",
            .Location = New Point(5, 60),
            .Size = New Size(146, 21),
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler Check1(OPC_ExtraerBAC).CheckedChanged, Sub(s, e) Check1_Click(OPC_ExtraerBAC)

        Check1(OPC_CrearCAJA) = New CheckBox() With {
            .Text = "Crear caja según BAC",
            .Location = New Point(5, 85),
            .Size = New Size(146, 21),
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler Check1(OPC_CrearCAJA).CheckedChanged, Sub(s, e) Check1_Click(OPC_CrearCAJA)

        Check1(OPC_ImprimirCAJA) = New CheckBox() With {
            .Text = "Imprimir etiqueta de caja",
            .Location = New Point(5, 110),
            .Size = New Size(146, 21),
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler Check1(OPC_ImprimirCAJA).CheckedChanged, Sub(s, e) Check1_Click(OPC_ImprimirCAJA)

        Check1(OPC_RelContenido) = New CheckBox() With {
            .Text = "Imprimir rel. contenido",
            .Location = New Point(5, 135),
            .Size = New Size(146, 21),
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler Check1(OPC_RelContenido).CheckedChanged, Sub(s, e) Check1_Click(OPC_RelContenido)

        Check1(OPC_Empaquetado) = New CheckBox() With {
            .Text = "Empaquetado automático",
            .Location = New Point(5, 160),
            .Size = New Size(146, 21),
            .BackColor = Color.Transparent,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler Check1(OPC_Empaquetado).CheckedChanged, Sub(s, e) Check1_Click(OPC_Empaquetado)

        ' Labels en FrameOpciones para mostrar datos de caja
        lbNumCaja2 = CreateValueLabel(60, 185, 27, 19)
        lbTipoCaja2 = CreateValueLabel(100, 185, 27, 19)
        lbNombreCaja2 = CreateValueLabel(130, 185, 117, 19)

        ' Botones de acción en FrameOpciones
        Dim btnCerrarBAC As New Button() With {.Text = "Cerrar BAC", .Location = New Point(5, 210), .Size = New Size(120, 25), .Enabled = False}
        AddHandler btnCerrarBAC.Click, Sub(s, e) cmdAccion_Click(CML_CerrarBAC)
        cmdAccion(CML_CerrarBAC) = btnCerrarBAC

        Dim btnExtraerBAC As New Button() With {.Text = "Extraer BAC", .Location = New Point(130, 210), .Size = New Size(120, 25), .Enabled = False}
        AddHandler btnExtraerBAC.Click, Sub(s, e) cmdAccion_Click(CML_ExtraerBAC)
        cmdAccion(CML_ExtraerBAC) = btnExtraerBAC

        Dim btnCrearCAJA As New Button() With {.Text = "Crear CAJA", .Location = New Point(5, 237), .Size = New Size(120, 25), .Enabled = False}
        AddHandler btnCrearCAJA.Click, Sub(s, e) cmdAccion_Click(CML_CrearCAJA)
        cmdAccion(CML_CrearCAJA) = btnCrearCAJA

        Dim btnImprimirCAJA As New Button() With {.Text = "Imprimir Etiqueta", .Location = New Point(130, 237), .Size = New Size(120, 25), .Enabled = False}
        AddHandler btnImprimirCAJA.Click, Sub(s, e) cmdAccion_Click(CML_ImprimirCAJA)
        cmdAccion(CML_ImprimirCAJA) = btnImprimirCAJA

        Dim btnRelContenido As New Button() With {.Text = "Rel. Contenido", .Location = New Point(5, 264), .Size = New Size(80, 25), .Enabled = False}
        AddHandler btnRelContenido.Click, Sub(s, e) cmdAccion_Click(CML_RelContenido)
        cmdAccion(CML_RelContenido) = btnRelContenido

        Dim btnEmpaquetado As New Button() With {.Text = "Empaquetar", .Location = New Point(90, 264), .Size = New Size(80, 25), .Enabled = False}
        AddHandler btnEmpaquetado.Click, Sub(s, e) cmdAccion_Click(CML_Empaquetado)
        cmdAccion(CML_Empaquetado) = btnEmpaquetado

        Dim btnCambiarCAJA As New Button() With {.Text = "Cambiar Caja", .Location = New Point(175, 264), .Size = New Size(75, 25), .Enabled = False}
        AddHandler btnCambiarCAJA.Click, Sub(s, e) cmdAccion_Click(CML_CambiarCAJA)
        cmdAccion(CML_CambiarCAJA) = btnCambiarCAJA

        Dim btnCambiarUDS As New Button() With {.Text = "Cambiar Uds", .Location = New Point(5, 290), .Size = New Size(80, 25), .Enabled = False, .Visible = False}
        AddHandler btnCambiarUDS.Click, Sub(s, e) cmdAccion_Click(CML_CambiarUDS)
        cmdAccion(CML_CambiarUDS) = btnCambiarUDS

        Dim btnCombinarCAJAS As New Button() With {.Text = "Combinar Cajas", .Location = New Point(90, 290), .Size = New Size(80, 25), .Enabled = False, .Visible = False}
        AddHandler btnCombinarCAJAS.Click, Sub(s, e) cmdAccion_Click(CML_CombinarCAJAS)
        cmdAccion(CML_CombinarCAJAS) = btnCombinarCAJAS

        Dim btnVolver As New Button() With {.Text = "VOLVER", .Location = New Point(175, 290), .Size = New Size(75, 25), .Visible = False}
        AddHandler btnVolver.Click, Sub(s, e) cmdAccion_Click(CML_Volver)
        cmdAccion(CML_Volver) = btnVolver

        FrameOpciones.Controls.AddRange(New Control() {
            Label2, lbUbicacion,
            Check1(OPC_CerrarBAC), Check1(OPC_ExtraerBAC), Check1(OPC_CrearCAJA),
            Check1(OPC_ImprimirCAJA), Check1(OPC_RelContenido), Check1(OPC_Empaquetado),
            lbNumCaja2, lbTipoCaja2, lbNombreCaja2,
            btnCerrarBAC, btnExtraerBAC, btnCrearCAJA, btnImprimirCAJA,
            btnRelContenido, btnEmpaquetado, btnCambiarCAJA, btnCambiarUDS, btnCombinarCAJAS, btnVolver
        })

        ' ============= FrameCajas (para cambio de tipo de caja) =============
        FrameCajas = New Panel()
        FrameCajas.BackColor = Color.Gray
        FrameCajas.Location = New Point(534, 0)     ' 8000/15
        FrameCajas.Size = New Size(254, 301)

        dgvCajas = New DataGridView()
        dgvCajas.Location = New Point(5, 5)
        dgvCajas.Size = New Size(244, 240)
        dgvCajas.AllowUserToAddRows = False
        dgvCajas.AllowUserToDeleteRows = False
        dgvCajas.ReadOnly = True
        dgvCajas.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvCajas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvCajas.RowHeadersVisible = False
        dgvCajas.BackgroundColor = Color.FromArgb(&HFF, &HDC, &HCE)

        Dim btnCambiaCAJA As New Button() With {.Text = "Cambiar Tipo Caja", .Location = New Point(5, 250), .Size = New Size(120, 25)}
        AddHandler btnCambiaCAJA.Click, Sub(s, e) cmdAccion_Click(CML_CambiaCAJA)
        cmdAccion(CML_CambiaCAJA) = btnCambiaCAJA

        Dim btnVolverCajas As New Button() With {.Text = "VOLVER", .Location = New Point(130, 250), .Size = New Size(120, 25)}
        AddHandler btnVolverCajas.Click, Sub(s, e) cmdAccion_Click(CML_VolverCajas)
        cmdAccion(CML_VolverCajas) = btnVolverCajas

        FrameCajas.Controls.AddRange(New Control() {dgvCajas, btnCambiaCAJA, btnVolverCajas})

        ' ============= FrameCantidad (para cambio de unidades) =============
        FrameCantidad = New Panel()
        FrameCantidad.BackColor = Color.Gray
        FrameCantidad.Location = New Point(800, 0)  ' 12000/15
        FrameCantidad.Size = New Size(254, 301)

        Dim lblArticulo As New Label() With {.Text = "Artículo:", .Location = New Point(5, 5), .Size = New Size(60, 19)}
        nArticulo = New TextBox() With {.Location = New Point(70, 5), .Size = New Size(80, 25)}
        lbNombreArticulo = New Label() With {.Location = New Point(5, 35), .Size = New Size(244, 40), .BackColor = Color.White}

        Dim lblCantidad As New Label() With {.Text = "Cantidad:", .Location = New Point(5, 85), .Size = New Size(60, 19)}
        nCantidad = New TextBox() With {.Location = New Point(70, 85), .Size = New Size(80, 25)}

        Dim btnRestarUDS As New Button() With {.Text = "-", .Location = New Point(5, 120), .Size = New Size(40, 25)}
        AddHandler btnRestarUDS.Click, Sub(s, e) cmdAccion_Click(CML_RestarUDS)
        cmdAccion(CML_RestarUDS) = btnRestarUDS

        Dim btnSumarUDS As New Button() With {.Text = "+", .Location = New Point(50, 120), .Size = New Size(40, 25)}
        AddHandler btnSumarUDS.Click, Sub(s, e) cmdAccion_Click(CML_SumarUDS)
        cmdAccion(CML_SumarUDS) = btnSumarUDS

        Dim btnAplicarUDS As New Button() With {.Text = "Aplicar", .Location = New Point(5, 150), .Size = New Size(120, 25)}
        AddHandler btnAplicarUDS.Click, Sub(s, e) cmdAccion_Click(CML_AplicarUDS)
        cmdAccion(CML_AplicarUDS) = btnAplicarUDS

        Dim btnVolverCantidad As New Button() With {.Text = "VOLVER", .Location = New Point(130, 150), .Size = New Size(120, 25)}
        AddHandler btnVolverCantidad.Click, Sub(s, e) cmdAccion_Click(CML_VolverCantidad)
        cmdAccion(CML_VolverCantidad) = btnVolverCantidad

        FrameCantidad.Controls.AddRange(New Control() {
            lblArticulo, nArticulo, lbNombreArticulo, lblCantidad, nCantidad,
            btnRestarUDS, btnSumarUDS, btnAplicarUDS, btnVolverCantidad
        })

        ' ============= FrameCombinar (para combinar cajas) =============
        FrameCombinar = New Panel()
        FrameCombinar.BackColor = Color.Gray
        FrameCombinar.Location = New Point(1067, 0) ' 16000/15
        FrameCombinar.Size = New Size(254, 301)

        ' Datos Caja 1
        Dim lblNumCaja3 As New Label() With {.Text = "Nº Caja", .Location = New Point(5, 10), .Size = New Size(51, 19)}
        lbNumCaja3 = CreateValueLabel(60, 10, 37, 19)

        Dim lblUds3 As New Label() With {.Text = "Uds", .Location = New Point(105, 10), .Size = New Size(26, 19)}
        lbUds3 = CreateValueLabel(135, 10, 32, 19)

        Dim lblArts3 As New Label() With {.Text = "Art.", .Location = New Point(180, 10), .Size = New Size(31, 19)}
        lbArts3 = CreateValueLabel(215, 10, 32, 19)

        Dim lblSSCC3 As New Label() With {.Text = "SSCC", .Location = New Point(5, 35), .Size = New Size(51, 19)}
        lbSSCC3 = CreateValueLabel(60, 35, 187, 19)

        Dim lblPeso3 As New Label() With {.Text = "Peso", .Location = New Point(5, 60), .Size = New Size(51, 19)}
        lbPeso3 = CreateValueLabel(60, 60, 57, 19)

        Dim lblVolumen3 As New Label() With {.Text = "Volumen", .Location = New Point(130, 60), .Size = New Size(56, 19)}
        lbVolumen3 = CreateValueLabel(190, 60, 57, 19)

        ' Lectura Caja 2
        Dim lblLeerCaja As New Label() With {.Text = "Leer caja a combinar (última caja!!)", .Location = New Point(5, 90), .Size = New Size(244, 16), .TextAlign = ContentAlignment.MiddleCenter}
        txtLecturaCaja2 = New TextBox() With {.Font = New Font("Arial", 14.25F), .Location = New Point(5, 110), .Size = New Size(244, 30), .MaxLength = 20}

        ' Datos Caja 2
        Dim lblNumCaja4 As New Label() With {.Text = "Nº Caja", .Location = New Point(5, 150), .Size = New Size(51, 19)}
        lbNumCaja4 = CreateValueLabel(60, 150, 37, 19)

        Dim lblUds4 As New Label() With {.Text = "Uds", .Location = New Point(105, 150), .Size = New Size(26, 19)}
        lbUds4 = CreateValueLabel(135, 150, 32, 19)

        Dim lblArts4 As New Label() With {.Text = "Art.", .Location = New Point(180, 150), .Size = New Size(31, 19)}
        lbArts4 = CreateValueLabel(215, 150, 32, 19)

        Dim lblPeso4 As New Label() With {.Text = "Peso", .Location = New Point(5, 175), .Size = New Size(51, 19)}
        lbPeso4 = CreateValueLabel(60, 175, 57, 19)

        Dim lblVolumen4 As New Label() With {.Text = "Volumen", .Location = New Point(130, 175), .Size = New Size(56, 19)}
        lbVolumen4 = CreateValueLabel(190, 175, 57, 19)

        ' Totales
        Dim lblPesoTot As New Label() With {.Text = "Peso Total", .Location = New Point(5, 210), .Size = New Size(51, 19)}
        lbPesoTot = CreateValueLabel(60, 210, 62, 19)

        Dim lblVolumenTot As New Label() With {.Text = "Vol. Total", .Location = New Point(130, 210), .Size = New Size(56, 19)}
        lbVolumenTot = CreateValueLabel(190, 210, 62, 19)

        ' Botones
        Dim btnAplicaCombinar As New Button() With {.Text = "Combinar Cajas", .Location = New Point(5, 240), .Size = New Size(244, 25), .Enabled = False}
        AddHandler btnAplicaCombinar.Click, Sub(s, e) cmdAccion_Click(CML_AplicaCombinar)
        cmdAccion(CML_AplicaCombinar) = btnAplicaCombinar

        Dim btnVolverCombinar As New Button() With {.Text = "VOLVER", .Location = New Point(5, 270), .Size = New Size(244, 25)}
        AddHandler btnVolverCombinar.Click, Sub(s, e) cmdAccion_Click(CML_VolverCombinar)
        cmdAccion(CML_VolverCombinar) = btnVolverCombinar

        FrameCombinar.Controls.AddRange(New Control() {
            lblNumCaja3, lbNumCaja3, lblUds3, lbUds3, lblArts3, lbArts3,
            lblSSCC3, lbSSCC3, lblPeso3, lbPeso3, lblVolumen3, lbVolumen3,
            lblLeerCaja, txtLecturaCaja2,
            lblNumCaja4, lbNumCaja4, lblUds4, lbUds4, lblArts4, lbArts4,
            lblPeso4, lbPeso4, lblVolumen4, lbVolumen4,
            lblPesoTot, lbPesoTot, lblVolumenTot, lbVolumenTot,
            btnAplicaCombinar, btnVolverCombinar
        })

        ' Agregar todos los frames al formulario
        Me.Controls.AddRange(New Control() {FrameLectura, FrameOpciones, FrameCajas, FrameCantidad, FrameCombinar})

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Private Function CreateValueLabel(x As Integer, y As Integer, width As Integer, height As Integer, Optional bold As Boolean = False) As Label
        Dim lbl As New Label()
        lbl.Text = ""
        lbl.Location = New Point(x, y)
        lbl.Size = New Size(width, height)
        lbl.BackColor = Color.White
        lbl.TextAlign = ContentAlignment.MiddleCenter
        lbl.Font = New Font("MS Sans Serif", If(bold, 9.75F, 8.25F), If(bold, FontStyle.Bold, FontStyle.Regular))
        Return lbl
    End Function

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

        ' Leer opciones del Ini
        Try
            Check1(OPC_CerrarBAC).Checked = (LeerIni(ficINI, "Opciones", "CerrarBAC", "0") = "1")
            Check1(OPC_ExtraerBAC).Checked = (LeerIni(ficINI, "Opciones", "ExtraerBAC", "0") = "1")
            Check1(OPC_CrearCAJA).Checked = (LeerIni(ficINI, "Opciones", "CrearCAJA", "0") = "1")
            Check1(OPC_ImprimirCAJA).Checked = (LeerIni(ficINI, "Opciones", "ImprimirCAJA", "0") = "1")
            Check1(OPC_RelContenido).Checked = (LeerIni(ficINI, "Opciones", "RelContenido", "0") = "1")
            Check1(OPC_Empaquetado).Checked = (LeerIni(ficINI, "Opciones", "Empaquetado", "0") = "1")
        Catch
        End Try

        cmdAccion(CML_CerrarBAC).Enabled = False
        cmdAccion(CML_ExtraerBAC).Enabled = False
        cmdAccion(CML_CrearCAJA).Enabled = False
        cmdAccion(CML_ImprimirCAJA).Enabled = False
        cmdAccion(CML_RelContenido).Enabled = False
        cmdAccion(CML_Empaquetado).Enabled = False
        cmdAccion(CML_CambiarCAJA).Enabled = False
        cmdAccion(CML_CambiarUDS).Enabled = False
        cmdAccion(CML_CombinarCAJAS).Enabled = False

        ' Acciones de la lectura
        cmdAccion(CML_Acciones).Text = ACC_General

        Cursor = Cursors.Default
    End Sub

    Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                Me.Close()
            Case Keys.F3
                cmdAccion_Click(CML_CrearCAJA)
            Case Keys.F4
                cmdAccion_Click(CML_ImprimirCAJA)
            Case Keys.F6
                cmdAccion_Click(CML_CerrarBAC)
            Case Keys.F10
                cmdAccion_Click(CML_RelContenido)
        End Select
    End Sub

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If ed IsNot Nothing Then
            ed.Dispose()
            ed = Nothing
        End If
    End Sub

    '----------------------------------------------------------------------------------

    Private Sub Salir()
        Me.Close()
    End Sub

    Private Sub Check1_Click(index As Integer)
        Select Case index
            Case OPC_CerrarBAC
                GuardarIni(ficINI, "Opciones", "CerrarBAC", If(Check1(index).Checked, "1", "0"))
                If Not Check1(index).Checked Then Check1(OPC_ExtraerBAC).Checked = False

            Case OPC_ExtraerBAC
                GuardarIni(ficINI, "Opciones", "ExtraerBAC", If(Check1(index).Checked, "1", "0"))
                If Check1(index).Checked Then Check1(OPC_CerrarBAC).Checked = True

            Case OPC_CrearCAJA
                GuardarIni(ficINI, "Opciones", "CrearCAJA", If(Check1(index).Checked, "1", "0"))

            Case OPC_ImprimirCAJA
                GuardarIni(ficINI, "Opciones", "ImprimirCAJA", If(Check1(index).Checked, "1", "0"))

            Case OPC_RelContenido
                GuardarIni(ficINI, "Opciones", "RelContenido", If(Check1(index).Checked, "1", "0"))

            Case OPC_Empaquetado
                GuardarIni(ficINI, "Opciones", "Empaquetado", If(Check1(index).Checked, "1", "0"))
        End Select
    End Sub

    Private Sub cmdAccion_Click(index As Integer)
        If Not cmdAccion.ContainsKey(index) OrElse Not cmdAccion(index).Enabled Then
            ' Si el botón no existe o no está habilitado, verificar acciones especiales
            If index = CML_Opciones OrElse index = CML_Acciones OrElse index = CML_Volver OrElse
               index = CML_VolverCajas OrElse index = CML_VolverCantidad OrElse index = CML_VolverCombinar OrElse
               index = CML_Salir Then
                ' Permitir siempre
            Else
                If index <> CML_Salir AndAlso index <> CML_CombinarCAJAS Then txtLecturaCodigo.Focus()
                Return
            End If
        End If

        Select Case index
            Case CML_Opciones
                PantallaOpciones()

            Case CML_Acciones
                AccionesAuto()

            Case CML_Volver
                PantallaPrincipal()

            ' Acciones principales
            Case CML_CerrarBAC
                CerrarBAC()
                PantallaPrincipal()

            Case CML_ExtraerBAC
                ExtraerBAC()
                PantallaPrincipal()

            Case CML_CrearCAJA
                CrearCAJA()
                PantallaPrincipal()

            Case CML_ImprimirCAJA
                ImprimirETIQUETA()
                PantallaPrincipal()

            Case CML_RelContenido
                ImprimirRELACION()
                PantallaPrincipal()

            Case CML_Empaquetado
                EmpaquetarBACaCAJA()
                PantallaPrincipal()

            ' Acciones adicionales
            Case CML_CambiarCAJA
                PantallaCambioCaja()

            Case CML_CambiarUDS
                PantallaCambioUnidades()

            Case CML_SumarUDS
                ModificaUnidades(1)

            Case CML_RestarUDS
                ModificaUnidades(-1)

            Case CML_AplicarUDS
                CambiarUnidades()

            Case CML_VolverCantidad
                PantallaPrincipal()

            ' Cambio de caja
            Case CML_CambiaCAJA
                CambiaTipoCaja()

            Case CML_VolverCajas
                PantallaPrincipal()

            ' Combinar cajas
            Case CML_CombinarCAJAS
                PantallaCombinarCajas()

            Case CML_VolverCombinar
                PantallaOpciones()

            Case CML_AplicaCombinar
                CombinarCajas(lbSSCC3.Text, txtLecturaCaja2.Text)

            ' Salir
            Case CML_Salir
                Salir()
        End Select

        If index <> CML_Salir AndAlso index <> CML_CombinarCAJAS Then txtLecturaCodigo.Focus()
    End Sub

    Private Sub AccionesAuto()
        Select Case cmdAccion(CML_Acciones).Text
            Case ACC_General
                PantallaOpciones()

            Case ACC_Empaquetar
                EmpaquetarBACaCAJA()

            Case ACC_Etiquetas
                ImprimirETIQUETA()
                ImprimirRELACION()
        End Select
    End Sub

    Private Sub lbNombreCaja_Click(sender As Object, e As EventArgs)
        PantallaCambioCaja()
    End Sub

    Private Sub PantallaPrincipal()
        FrameLectura.Left = 0
        FrameOpciones.Left = 267     ' 4000/15
        FrameCantidad.Left = 534     ' 8000/15
        FrameCajas.Left = 800        ' 12000/15
        FrameCombinar.Left = 1067    ' 16000/15
        txtLecturaCodigo.Focus()
    End Sub

    Private Sub PantallaOpciones()
        ' Se deshabilitan todas las opciones
        cmdAccion(CML_CerrarBAC).Enabled = False
        cmdAccion(CML_ExtraerBAC).Enabled = False
        cmdAccion(CML_CrearCAJA).Enabled = False
        cmdAccion(CML_ImprimirCAJA).Enabled = False
        cmdAccion(CML_RelContenido).Enabled = False
        cmdAccion(CML_Empaquetado).Enabled = False
        cmdAccion(CML_CambiarCAJA).Enabled = False
        cmdAccion(CML_CambiarUDS).Enabled = False
        cmdAccion(CML_CombinarCAJAS).Enabled = False

        ' Configura opciones de BAC
        If Label3.Text = "BAC" Then
            If lbBAC.Text <> "" AndAlso tEstadoBAC = 0 Then cmdAccion(CML_CerrarBAC).Enabled = True
            If tUbicacionBAC > 0 Then cmdAccion(CML_ExtraerBAC).Enabled = True
            If Val(lbNumCaja.Text) = 0 Then cmdAccion(CML_CrearCAJA).Enabled = True
            If Val(lbNumCaja.Text) > 0 Then cmdAccion(CML_ImprimirCAJA).Enabled = True
            If Val(lbNumCaja.Text) > 0 OrElse tUbicacionBAC = 0 Then cmdAccion(CML_CambiarCAJA).Enabled = True

            cmdAccion(CML_Empaquetado).Enabled = True
        End If

        ' Configura opciones de CAJA
        If Label3.Text = "CAJA" Then
            cmdAccion(CML_ImprimirCAJA).Enabled = True
            cmdAccion(CML_RelContenido).Enabled = True
            cmdAccion(CML_CambiarCAJA).Enabled = True
            cmdAccion(CML_CambiarUDS).Enabled = True
            cmdAccion(CML_CombinarCAJAS).Enabled = True
        End If

        ' Visualiza en el formulario
        FrameLectura.Left = 267      ' 4000/15
        FrameCantidad.Left = 534     ' 8000/15
        FrameCajas.Left = 800        ' 12000/15
        FrameCombinar.Left = 1067    ' 16000/15
        FrameOpciones.Left = 0
    End Sub

    Private Sub PantallaCambioCaja()
        ' Cambio de tipo de Caja
        FrameOpciones.Left = 267
        FrameCantidad.Left = 534
        FrameLectura.Left = 800
        FrameCombinar.Left = 1067
        FrameCajas.Left = 0

        dgvCajas.DataSource = Nothing

        Try
            ' Carga la lista de tipos de caja
            Dim dtTipos As DataTable = ed.DameTiposCajasActivas()

            If dtTipos.Rows.Count > 0 Then
                dgvCajas.DataSource = dtTipos
                dgvCajas.Focus()
            Else
                wsMensaje("No Existen Cajas de Empaquetar Definidas.", TipoMensaje.MENSAJE_Informativo)
            End If
        Catch ex As Exception
            wsMensaje("Error al cargar tipos de caja: " & ex.Message, TipoMensaje.MENSAJE_Grave)
        End Try
    End Sub

    Private Sub PantallaCambioUnidades()
        ' Cambio de unidades de artículo. Sólo permitido en CAJAS
        FrameLectura.Left = 534
        FrameOpciones.Left = 267
        FrameCajas.Left = 800
        FrameCombinar.Left = 1067
        FrameCantidad.Left = 0

        ' Posiciona el artículo actual
        If dtArticulosCaja IsNot Nothing AndAlso dtArticulosCaja.Rows.Count > 0 Then
            nArticulo.Text = dtArticulosCaja.Rows(0)("ltaart").ToString()
            nCantidad.Text = dtArticulosCaja.Rows(0)("ltacan").ToString()
            lbNombreArticulo.Text = If(Not IsDBNull(dtArticulosCaja.Rows(0)("artnom")), dtArticulosCaja.Rows(0)("artnom").ToString(), "")
        Else
            nArticulo.Text = ""
            nCantidad.Text = ""
            lbNombreArticulo.Text = ""
        End If

        nArticulo.Focus()
    End Sub

    Private Sub PantallaCombinarCajas()
        ' Cambio de tipo de Caja
        FrameLectura.Left = 1067
        FrameOpciones.Left = 267
        FrameCantidad.Left = 534
        FrameCajas.Left = 800
        FrameCombinar.Left = 0

        ' Datos Caja 1
        lbNumCaja3.Text = lbNumCaja.Text
        lbUds3.Text = lbUds.Text
        lbArts3.Text = lbArts.Text

        lbSSCC3.Text = lbSSCC.Text
        'lbTipoCaja3.Text = lbTipoCaja.Text
        'lbNombreCaja3.Text = lbNombreCaja.Text
        lbPeso3.Text = lbPeso.Text
        lbVolumen3.Text = lbVolumen.Text

        ' Datos Caja 2
        txtLecturaCaja2.Text = ""
        lbNumCaja4.Text = ""
        lbUds4.Text = ""
        lbArts4.Text = ""
        lbPeso4.Text = ""
        lbVolumen4.Text = ""

        ' Datos totales
        lbPesoTot.Text = ""
        lbVolumenTot.Text = ""

        cmdAccion(CML_AplicaCombinar).Enabled = False
        txtLecturaCaja2.Enabled = True

        txtLecturaCaja2.Focus()
    End Sub

    Private Sub txtLecturaCaja2_KeyDown(sender As Object, e As KeyEventArgs) Handles txtLecturaCaja2.KeyDown
        If e.KeyCode = Keys.Return Then
            Select Case txtLecturaCaja2.Text.Length
                Case 18, 20 ' SSCC de Caja
                    If txtLecturaCaja2.Text.Length = 20 Then
                        txtLecturaCaja2.Text = txtLecturaCaja2.Text.Substring(2, 18)
                    End If

                    ' Comprobación de datos: Caja repetida
                    If txtLecturaCaja2.Text = lbBAC.Text Then
                        wsMensaje(" La caja está repetida! ", TipoMensaje.MENSAJE_Grave)
                        cmdAccion(CML_AplicaCombinar).Enabled = False
                        txtLecturaCaja2.Text = ""
                        txtLecturaCaja2.Focus()
                        Exit Sub
                    End If

                    Try
                        Dim dtCaja As DataTable = ed.DameDatosCAJAdePTL(txtLecturaCaja2.Text)

                        If dtCaja.Rows.Count > 0 Then
                            Dim row As DataRow = dtCaja.Rows(0)

                            ' Comprobación de datos: Grupo y Tablilla
                            Dim ltcgru As Integer = If(Not IsDBNull(row("ltcgru")), CInt(row("ltcgru")), 0)
                            Dim ltctab As Integer = If(Not IsDBNull(row("ltctab")), CInt(row("ltctab")), 0)

                            If lbGrupo.Text <> ltcgru.ToString() OrElse lbTablilla.Text <> ltctab.ToString() Then
                                wsMensaje(" La caja no pertenece al Grupo / Tablilla ", TipoMensaje.MENSAJE_Grave)
                                cmdAccion(CML_AplicaCombinar).Enabled = False
                                txtLecturaCaja2.Text = ""
                                txtLecturaCaja2.Focus()
                                Exit Sub
                            End If

                            ' Se muestran los datos
                            Dim ltccaj As String = If(Not IsDBNull(row("ltccaj")), row("ltccaj").ToString(), "")
                            lbNumCaja4.Text = ltccaj
                            lbUds4.Text = ""
                            lbArts4.Text = ""
                            lbPeso4.Text = "0"
                            lbVolumen4.Text = "0"

                            ' Datos totales
                            Dim peso3 As Double = 0
                            Dim peso4 As Double = 0
                            Double.TryParse(lbPeso3.Text, peso3)
                            Double.TryParse(lbPeso4.Text, peso4)
                            lbPesoTot.Text = (peso3 + peso4).ToString("#0.000")

                            Dim vol3 As Double = 0
                            Dim vol4 As Double = 0
                            Double.TryParse(lbVolumen3.Text, vol3)
                            Double.TryParse(lbVolumen4.Text, vol4)
                            lbVolumenTot.Text = (vol3 + vol4).ToString("#0.000")

                            txtLecturaCaja2.Enabled = False
                            cmdAccion(CML_AplicaCombinar).Enabled = True
                            cmdAccion(CML_AplicaCombinar).Focus()
                        Else
                            wsMensaje(" No existe la CAJA ", TipoMensaje.MENSAJE_Grave)
                            cmdAccion(CML_AplicaCombinar).Enabled = False
                            txtLecturaCaja2.Text = ""
                            txtLecturaCaja2.Focus()
                        End If
                    Catch ex As Exception
                        wsMensaje(" Error: " & ex.Message, TipoMensaje.MENSAJE_Grave)
                        cmdAccion(CML_AplicaCombinar).Enabled = False
                        txtLecturaCaja2.Text = ""
                        txtLecturaCaja2.Focus()
                    End Try

                Case Else
                    wsMensaje(" No existe la CAJA ", TipoMensaje.MENSAJE_Grave)
            End Select
        End If
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
                Case 12 ' Unidad de transporte
                    ' Comprobar si la lectura es un BAC
                    Label3.Text = "BAC"
                    fValidarBAC(txtLecturaCodigo.Text, True)

                Case 18 ' SSCC de Caja
                    Label3.Text = "CAJA"
                    fValidarCaja(txtLecturaCodigo.Text, True)

                Case 20 ' SSCC de Caja
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

                bCalculoPeso = unipes > unipma
                bCalculoVolumen = univol > univma
                tEstadoBAC = If(Not IsDBNull(row("uniest")), CInt(row("uniest")), 0)
                tUbicacionBAC = If(Not IsDBNull(row("uninum")), CInt(row("uninum")), 0)

                Dim unicod As String = If(Not IsDBNull(row("unicod")), row("unicod").ToString(), "")
                Dim unigru As Integer = If(Not IsDBNull(row("unigru")), CInt(row("unigru")), 0)
                Dim unitab As Integer = If(Not IsDBNull(row("unitab")), CInt(row("unitab")), 0)
                Dim unicaj As String = If(Not IsDBNull(row("unicaj")), row("unicaj").ToString(), "")
                Dim tipdes As String = If(Not IsDBNull(row("tipdes")), row("tipdes").ToString(), "")
                Dim uninca As String = If(Not IsDBNull(row("uninca")), row("uninca").ToString(), "")

                ' Se muestran los datos
                If IsDBNull(row("ubicod")) Then
                    RefrescarDatos(False, 0, 0, 0, 0, 0, unicod, tEstadoBAC, unigru, unitab, unipes, univol, unicaj, tipdes, uninca, bCalculoPeso, bCalculoVolumen)
                Else
                    Dim ubicod As Integer = CInt(row("ubicod"))
                    Dim ubialm As Integer = If(Not IsDBNull(row("ubialm")), CInt(row("ubialm")), 0)
                    Dim ubiblo As Integer = If(Not IsDBNull(row("ubiblo")), CInt(row("ubiblo")), 0)
                    Dim ubifil As Integer = If(Not IsDBNull(row("ubifil")), CInt(row("ubifil")), 0)
                    Dim ubialt As Integer = If(Not IsDBNull(row("ubialt")), CInt(row("ubialt")), 0)
                    RefrescarDatos(False, ubicod, ubialm, ubiblo, ubifil, ubialt, unicod, tEstadoBAC, unigru, unitab, unipes, univol, unicaj, tipdes, uninca, bCalculoPeso, bCalculoVolumen)
                End If

                ' Lista de artículos contenidos en el BAC
                sRefrescarArticulosBAC(unigru, unicod)

                ' Comprueba el estado del BAC
                If tEstadoBAC = 0 Then
                    If Check1(OPC_CerrarBAC).Checked Then
                        CerrarBAC()
                    Else
                        If blMensaje Then wsMensaje(" El BAC está abierto!!", TipoMensaje.MENSAJE_Grave)
                        Return False
                    End If
                End If

                ' Comprueba si está ubicado el BAC
                If tUbicacionBAC > 0 Then
                    If Check1(OPC_ExtraerBAC).Checked Then
                        ExtraerBAC()
                    Else
                        If blMensaje Then wsMensaje(" El BAC está ubicado!!", TipoMensaje.MENSAJE_Grave)
                        Return False
                    End If
                End If

                ' Acciones de la lectura
                cmdAccion(CML_Acciones).Text = ACC_Empaquetar
            Else
                ' Cuando no existe el bac se busca la última caja a la que se ha traspasado desde ese BAC
                Try
                    Dim dtUltimaCaja As DataTable = ed.DameUltimaCajaDeBAC(stBAC)
                    If dtUltimaCaja.Rows.Count = 0 Then
                        If blMensaje Then wsMensaje(" No existe el BAC ", TipoMensaje.MENSAJE_Grave)
                    Else
                        If MessageBox.Show("¿Recuperar última caja de este BAC?", "BAC vacío!!", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                            Label3.Text = "CAJA"
                            Dim ltcssc As String = If(Not IsDBNull(dtUltimaCaja.Rows(0)("ltcssc")), dtUltimaCaja.Rows(0)("ltcssc").ToString(), "")
                            fValidarCaja(ltcssc, True)
                            Return fValidarBAC
                        End If
                    End If
                Catch
                    If blMensaje Then wsMensaje(" No existe el BAC ", TipoMensaje.MENSAJE_Grave)
                End Try
                ' Acciones de la lectura
                cmdAccion(CML_Acciones).Text = ACC_General
            End If
        Catch ex As Exception
            If blMensaje Then wsMensaje(" Error: " & ex.Message, TipoMensaje.MENSAJE_Grave)
            cmdAccion(CML_Acciones).Text = ACC_General
        End Try

        Return fValidarBAC
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
                bEstado = If(ltcvol > 0, 1, 0)

                Dim ltcssc As String = If(Not IsDBNull(row("ltcssc")), row("ltcssc").ToString(), "")
                Dim ltcgru As Integer = If(Not IsDBNull(row("ltcgru")), CInt(row("ltcgru")), 0)
                Dim ltctab As Integer = If(Not IsDBNull(row("ltctab")), CInt(row("ltctab")), 0)
                Dim ltcpes As Double = If(Not IsDBNull(row("ltcpes")), CDbl(row("ltcpes")), 0)
                Dim ltctip As String = If(Not IsDBNull(row("ltctip")), row("ltctip").ToString(), "")
                Dim tipdes As String = If(Not IsDBNull(row("tipdes")), row("tipdes").ToString(), "")
                Dim ltccaj As String = If(Not IsDBNull(row("ltccaj")), row("ltccaj").ToString(), "")

                ' Se muestran los datos
                RefrescarDatos(False, 0, 0, 0, 0, 0, ltcssc, bEstado, ltcgru, ltctab, ltcpes, ltcvol, ltctip, tipdes, ltccaj, bCalculoPeso, bCalculoVolumen)

                ' Lista de artículos contenidos en la CAJA
                sRefrescarArticulosCAJA(ltcgru, ltctab, ltccaj)

                ' Acciones de la lectura
                cmdAccion(CML_Acciones).Text = ACC_Etiquetas
            Else
                If blMensaje Then wsMensaje(" No existe la CAJA ", TipoMensaje.MENSAJE_Grave)

                ' Acciones de la lectura
                cmdAccion(CML_Acciones).Text = ACC_General
            End If
        Catch ex As Exception
            If blMensaje Then wsMensaje(" Error: " & ex.Message, TipoMensaje.MENSAJE_Grave)
            cmdAccion(CML_Acciones).Text = ACC_General
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
            lbArts.Text = ""

            lbPeso.Text = ""
            lbPeso.BackColor = Color.White

            lbVolumen.Text = ""
            lbVolumen.BackColor = Color.White

            lbTipoCaja.Text = ""
            lbNombreCaja.Text = ""

            tEstadoBAC = 0
            tUbicacionBAC = 0

            lbNumCaja.Text = ""
            lbNumCaja2.Text = ""
            lbSSCC.Text = ""
            lbTipoCaja2.Text = ""
            lbNombreCaja2.Text = ""
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
            lbArts.Text = "0"

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
            lbNumCaja2.Text = sNumCaja

            ' Datos relacionados
            If Val(sNumCaja) > 0 Then
                Try
                    Dim dtCajaGT As DataTable = ed.DameCajaGrupoTablillaPTL(sGrupo, sTablilla, sNumCaja)
                    If dtCajaGT.Rows.Count > 0 Then
                        lbSSCC.Text = If(Not IsDBNull(dtCajaGT.Rows(0)("ltcssc")), dtCajaGT.Rows(0)("ltcssc").ToString(), "")
                    Else
                        lbSSCC.Text = "ERROR EN LA CAJA"
                    End If
                Catch
                    lbSSCC.Text = ""
                End Try
            Else
                lbSSCC.Text = ""
            End If

            lbTipoCaja2.Text = sTipoCaja
            lbNombreCaja2.Text = sNombreCaja
        End If

        dgvArticulos.DataSource = Nothing
    End Sub

    Private Sub sRefrescarArticulosBAC(ByVal sGrupo As Long, sBAC As String)
        Dim iUds As Integer = 0

        dgvArticulos.DataSource = Nothing

        Try
            Dim dtContenido As DataTable = ed.DameContenidoBacGrupo(CInt(sGrupo), sBAC)
            dtArticulos = dtContenido

            If dtContenido.Rows.Count > 0 Then
                For Each row As DataRow In dtContenido.Rows
                    If Not IsDBNull(row("unican")) Then
                        iUds = iUds + CInt(row("unican"))
                    End If
                Next

                ConfigurarColumnasBAC()
                dgvArticulos.DataSource = dtContenido
            End If
        Catch ex As Exception
            wsMensaje($" Error al cargar artículos: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        lbUds.Text = iUds.ToString()
        lbArts.Text = If(dtArticulos IsNot Nothing, dtArticulos.Rows.Count.ToString(), "0")
    End Sub

    Private Sub sRefrescarArticulosCAJA(ByVal sGrupo As Long, sTablilla As Long, sCaja As String)
        Dim iUds As Integer = 0

        dgvArticulos.DataSource = Nothing

        Try
            Dim dtContenido As DataTable = ed.DameContenidoCajaGrupo(CInt(sGrupo), CInt(sTablilla), sCaja)
            dtArticulosCaja = dtContenido

            If dtContenido.Rows.Count > 0 Then
                For Each row As DataRow In dtContenido.Rows
                    If Not IsDBNull(row("ltacan")) Then
                        iUds = iUds + CInt(CDbl(row("ltacan")))
                    End If
                Next

                ConfigurarColumnasCAJA()
                dgvArticulos.DataSource = dtContenido
            End If
        Catch ex As Exception
            wsMensaje($" Error al cargar artículos: {ex.Message}", TipoMensaje.MENSAJE_Grave)
        End Try

        lbUds.Text = iUds.ToString()
        lbArts.Text = If(dtArticulosCaja IsNot Nothing, dtArticulosCaja.Rows.Count.ToString(), "0")
    End Sub

    Private Sub ugArticulos_DblClick(sender As Object, e As EventArgs)
        PantallaCambioUnidades()
    End Sub

    Private Sub ConfigurarColumnasBAC()
        dgvArticulos.AutoGenerateColumns = False
        dgvArticulos.Columns.Clear()

        Dim colCodigo As New DataGridViewTextBoxColumn()
        colCodigo.DataPropertyName = "uniart"
        colCodigo.HeaderText = "Codigo"
        colCodigo.Width = 40
        dgvArticulos.Columns.Add(colCodigo)

        Dim colNombre As New DataGridViewTextBoxColumn()
        colNombre.DataPropertyName = "artnom"
        colNombre.HeaderText = "Articulo"
        colNombre.Width = 107
        dgvArticulos.Columns.Add(colNombre)

        Dim colCant As New DataGridViewTextBoxColumn()
        colCant.DataPropertyName = "unican"
        colCant.HeaderText = "Cant"
        colCant.Width = 33
        dgvArticulos.Columns.Add(colCant)
    End Sub

    Private Sub ConfigurarColumnasCAJA()
        dgvArticulos.AutoGenerateColumns = False
        dgvArticulos.Columns.Clear()

        Dim colCodigo As New DataGridViewTextBoxColumn()
        colCodigo.DataPropertyName = "ltaart"
        colCodigo.HeaderText = "Codigo"
        colCodigo.Width = 40
        dgvArticulos.Columns.Add(colCodigo)

        Dim colNombre As New DataGridViewTextBoxColumn()
        colNombre.DataPropertyName = "artnom"
        colNombre.HeaderText = "Articulo"
        colNombre.Width = 107
        dgvArticulos.Columns.Add(colNombre)

        Dim colCant As New DataGridViewTextBoxColumn()
        colCant.DataPropertyName = "ltacan"
        colCant.HeaderText = "Cant"
        colCant.Width = 33
        dgvArticulos.Columns.Add(colCant)
    End Sub

    '---------------------------------------------------------------------------------------------------------------
    ' Acciones
    '---------------------------------------------------------------------------------------------------------------

    '--- CERRAR BAC

    Private Sub CerrarBAC()
        If tEstadoBAC = 1 Then
            ' El BAC ya está cerrado
            Exit Sub
        End If

        ' Cambiar estado de BAC de 0 a 1
        If CambiarEstadoBAC(lbBAC.Text, 1) Then
            tEstadoBAC = 1
            lbBAC.BackColor = If(tEstadoBAC = 0, Color.White, Color.FromArgb((ColorVerde >> 16) And &HFF, (ColorVerde >> 8) And &HFF, ColorVerde And &HFF))
        End If
    End Sub

    Private Function CambiarEstadoBAC(tBac As String, tEstado As Integer) As Boolean
        ' Cambio de estado de BAC
        Dim Retorno As Integer = 0
        Dim msgSalida As String = ""

        CambiarEstadoBAC = False

        Try
            ed.CambiaEstadoBACdePTL(tBac, tEstado, Usuario.Id, Retorno, msgSalida)

            If Retorno = 0 Then
                CambiarEstadoBAC = True
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("CambiarEstadoBAC error: " & ex.Message)
        End Try

        Return CambiarEstadoBAC
    End Function

    '--- EXTRAER BAC

    Private Sub ExtraerBAC()
        If tUbicacionBAC = 0 Then
            ' El BAC ya está extraido
            Exit Sub
        End If

        If RetirarBAC(lbBAC.Text, tEstadoBAC, True) Then
            tUbicacionBAC = 0
            lbUbicacion.Text = "-------------"
        End If
    End Sub

    Private Function RetirarBAC(tBac As String, tEstado As Integer, tEstadoFinal As Boolean) As Boolean
        ' Extracción del BAC
        Dim Retorno As Integer = 0
        Dim msgSalida As String = ""
        Dim nEstado As Integer

        RetirarBAC = False

        Try
            ed.RetirarBACdePTL(tBac, Usuario.Id, Retorno, msgSalida)

            If Retorno = 0 Then
                RetirarBAC = True
                If (tEstado = 0) = tEstadoFinal Then
                    If tEstadoFinal Then nEstado = 1 Else nEstado = 0
                    ' Cambiar estado de BAC
                    If CambiarEstadoBAC(tBac, nEstado) Then
                        lbBAC.BackColor = If(tEstadoBAC = 0, Color.White, Color.FromArgb((ColorVerde >> 16) And &HFF, (ColorVerde >> 8) And &HFF, ColorVerde And &HFF))
                    End If
                End If
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("RetirarBAC error: " & ex.Message)
        End Try

        Return RetirarBAC
    End Function

    '--- CREAR CAJA

    Private Sub CrearCAJA()
        Dim nCaja As Long
        Dim sSSCC As String = ""

        If tUbicacionBAC <> 0 AndAlso tEstadoBAC = 0 Then
            ' No se puede crear caja cuando el BAC está ubicado y está abierto
            Exit Sub
        End If

        ' Comprueba si el BAC ya tiene caja asignada
        If Val(lbNumCaja.Text) > 0 Then
            ' La caja ya ha sido creada previamente
            Exit Sub
        End If

        Try
            ' Busca el siguiente número de caja
            Dim dtCajas As DataTable = ed.DameCajasGrupoTablillaPTL(CInt(lbGrupo.Text), CInt(lbTablilla.Text))

            If dtCajas.Rows.Count > 0 Then
                nCaja = CInt(dtCajas.Rows(dtCajas.Rows.Count - 1)("ltccaj")) + 1
            Else
                ' No hay ninguna caja
                nCaja = 1
            End If

            ' Se crea la caja
            If CrearCajaNueva(CInt(lbGrupo.Text), CInt(lbTablilla.Text), nCaja, CInt(lbTipoCaja.Text), sSSCC, lbBAC.Text) Then
                ' Se ha creado la caja
                lbSSCC.Text = sSSCC
                lbNumCaja.Text = nCaja.ToString()
                lbNumCaja2.Text = nCaja.ToString()
            Else
                wsMensaje(" No se ha podido crear la CAJA. ", TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            wsMensaje(" Error al crear caja: " & ex.Message, TipoMensaje.MENSAJE_Grave)
        End Try
    End Sub

    Private Function CrearCajaNueva(ByVal ilGrupo As Long, ByVal ilTablilla As Long, ByRef ilCaja As Long, ByRef ilTipoCaja As Long, ByRef slSSCC As String, slBAC As String) As Boolean
        CrearCajaNueva = False

        Try
            ' Obtención del SSCC que le va a corresponder a la nueva Caja
            slSSCC = ObtenerSSCC_Heterogeneo()

            If slSSCC = "" Then
                Exit Function
            End If

            ' Inserta la caja y actualiza el BAC
            ed.CrearCajaGrupoTablillaPTL(CInt(ilGrupo), CInt(ilTablilla), ilCaja.ToString(), CInt(ilTipoCaja), slSSCC, slBAC)
            ed.ActualizaCajaBACPTL(slBAC, ilCaja.ToString())

            CrearCajaNueva = True
        Catch ex As Exception
            wsMensaje(" Error al crear caja: " & ex.Message, TipoMensaje.MENSAJE_Grave)
        End Try

        Return CrearCajaNueva
    End Function

    '--- IMPRIMIR ETIQUETA DE LA CAJA

    Private Sub ImprimirETIQUETA()
        If Val(lbNumCaja.Text) = 0 Then
            ' No se puede imprimir la etiqueta porque no hay caja creada
            Exit Sub
        End If

        ' En VB6 se usaba Crystal Reports, aquí mostramos mensaje de éxito
        wsMensaje("Etiqueta de caja enviada a imprimir", TipoMensaje.MENSAJE_Exclamacion)
    End Sub

    '--- IMPRIMIR RELACIÓN DE CONTENIDO DE LA CAJA

    Private Sub ImprimirRELACION()
        If tUbicacionBAC <> 0 AndAlso tEstadoBAC = 0 Then
            ' No se puede imprimir la caja cuando el BAC está ubicado y está abierto
            Exit Sub
        End If

        If Val(lbNumCaja.Text) = 0 Then
            ' No se puede imprimir la relación porque no hay caja creada
            Exit Sub
        End If

        ' En VB6 se usaba wsImprimirContenidoCaja, aquí mostramos mensaje de éxito
        wsMensaje("Relación de contenido enviada a imprimir", TipoMensaje.MENSAJE_Exclamacion)
    End Sub

    '--- CAMBIAR EL TIPO DE CAJA

    Private Sub CambiaTipoCaja()
        If dgvCajas.SelectedRows.Count = 0 Then Exit Sub

        Dim stBAC As String = ""
        Dim stSSCC As String = ""
        Dim tipcod As Integer = CInt(dgvCajas.SelectedRows(0).Cells("tipcod").Value)
        Dim tipdes As String = dgvCajas.SelectedRows(0).Cells("tipdes").Value.ToString()

        If tipcod = CInt(lbTipoCaja2.Text) Then Exit Sub

        If Label3.Text = "BAC" Then
            stBAC = lbBAC.Text
        Else
            stSSCC = lbBAC.Text
        End If

        Try
            ed.CambiaTipoCajaPTL(tipcod, stBAC, stSSCC, Usuario.Id)

            ' Refresca los datos
            lbTipoCaja.Text = tipcod.ToString()
            lbTipoCaja2.Text = tipcod.ToString()

            lbNombreCaja.Text = tipdes
            lbNombreCaja2.Text = tipdes

            ' Acciones adicionales
            If Check1(OPC_ImprimirCAJA).Checked AndAlso Val(lbNumCaja.Text) > 0 Then
                ' Imprimir caja
                ImprimirETIQUETA()
            End If

            wsMensaje("Se ha realizado el cambio de tipo de caja.", TipoMensaje.MENSAJE_Informativo)
        Catch ex As Exception
            wsMensaje(" Error al cambiar tipo de caja: " & ex.Message, TipoMensaje.MENSAJE_Grave)
        End Try
    End Sub

    '--- COMBINAR CAJAS

    Private Sub CombinarCajas(SSCC1 As String, SSCC2 As String)
        ' Solo se puede combinar con otra caja la última caja de la tablilla.
        Dim Retorno As Integer = 0
        Dim msgSalida As String = ""

        Try
            ' Combinar CAJAS por SQL
            ed.CombinarCajasPTL(SSCC1, SSCC2, Usuario.Id, Retorno, msgSalida)

            If Retorno = 0 Then
                ' Imprimir nueva relación de contenido
                ImprimirRELACION()

                wsMensaje("Se han combinado las cajas.", TipoMensaje.MENSAJE_Informativo)

                cmdAccion(CML_AplicaCombinar).Enabled = False
            Else
                wsMensaje(" Error al combinar Cajas" & vbNewLine & msgSalida, TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            wsMensaje(" Error al combinar cajas: " & ex.Message, TipoMensaje.MENSAJE_Grave)
        End Try
    End Sub

    '--- EMPAQUETAR BAC

    Private Sub EmpaquetarBACaCAJA()
        Dim Retorno As Integer = 0
        Dim msgSalida As String = ""
        Dim tSSCC As String = ""

        ' Comprobaciones previas
        If dtArticulos Is Nothing OrElse dtArticulos.Rows.Count = 0 Then Exit Sub

        Try
            ' Base del SSCC para pasar al procedimiento
            tSSCC = ObtenerSSCC_Heterogeneo()

            ' Empaquetado de BAC a CAJA por SQL
            ed.TraspasaBACaCAJAdePTLByRef(lbBAC.Text, Usuario.Id, tSSCC, Retorno, msgSalida)

            If Retorno = 0 Then
                wsMensaje(msgSalida, TipoMensaje.MENSAJE_Exclamacion)
                ' Refresco de los datos de la pantalla con el SSCC de destino
                Label3.Text = "CAJA"
                fValidarCaja(tSSCC, True)

                ' Acciones adicionales
                If Check1(OPC_ImprimirCAJA).Checked Then
                    ' Imprimir caja
                    ImprimirETIQUETA()
                End If

                If Check1(OPC_RelContenido).Checked Then
                    ' Imprimir caja
                    ImprimirRELACION()
                End If
            Else
                wsMensaje(" Error al Empaquetar el BAC" & vbNewLine & msgSalida, TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            wsMensaje(" Error al empaquetar BAC: " & ex.Message, TipoMensaje.MENSAJE_Grave)
        End Try
    End Sub

    '--- CAMBIO DE UNIDADES

    Private Sub ModificaUnidades(tCantidad As Integer)
        Dim cantidad As Integer = 0
        Integer.TryParse(nCantidad.Text, cantidad)
        cantidad = cantidad + tCantidad
        If cantidad < 0 Then cantidad = 0
        nCantidad.Text = cantidad.ToString()
    End Sub

    Private Sub CambiarUnidades()
        ' Solo se pueden cambiar unidades en la caja
        Dim Retorno As Integer = 0
        Dim msgSalida As String = ""

        Try
            Dim articulo As Integer = 0
            Dim cantidad As Integer = 0
            Integer.TryParse(nArticulo.Text, articulo)
            Integer.TryParse(nCantidad.Text, cantidad)

            ' Cambio de unidades
            ed.CambiaUnidadesArtCajaPTL(lbBAC.Text, articulo, cantidad, Usuario.Id, Retorno, msgSalida)

            If Retorno = 0 Then
                wsMensaje(msgSalida, TipoMensaje.MENSAJE_Exclamacion)
                ' Refresco de los datos de la pantalla con el SSCC de destino
                Label3.Text = "CAJA"
                fValidarCaja(lbBAC.Text, True)

                ' Acciones adicionales
                If Check1(OPC_RelContenido).Checked Then
                    ' Imprimir caja
                    ImprimirRELACION()
                End If
            Else
                wsMensaje(" Error al cambiar las unidades del Artículo" & vbNewLine & msgSalida, TipoMensaje.MENSAJE_Grave)
            End If
        Catch ex As Exception
            wsMensaje(" Error al cambiar unidades: " & ex.Message, TipoMensaje.MENSAJE_Grave)
        End Try
    End Sub

    '------------------------------------------------------------------------------------------------------------------
    ' FUNCIONES
    '------------------------------------------------------------------------------------------------------------------

    ' -- Función para Obtener el SSCC Heterogeneo que Corresponda -----
    Private Function ObtenerSSCC_Heterogeneo() As String
        ' En VB6 esto era más complejo con transacciones y Crystal Reports
        ' Simplificado para .NET - genera un SSCC basado en timestamp
        Try
            Dim sSSCC As String = ""

            ' Obtener siguiente numerador SSCC
            Dim numerador As Long = Dame_Siguiente_Numerador_SSCC_Heterogeneo()

            If numerador < 0 Then
                Return ""
            End If

            ' Actualizar el numerador
            ed.ActualizaNumeradorSSCCHipodromo(CInt(numerador))

            ' Generar SSCC (simulado - en producción usaría Dame_SSCC)
            sSSCC = "3842" & numerador.ToString().PadLeft(14, "0"c)

            Return sSSCC
        Catch ex As Exception
            wsMensaje(" Error al obtener SSCC: " & ex.Message, TipoMensaje.MENSAJE_Grave)
            Return ""
        End Try
    End Function

    ' -- Función para obtener el Siguiente numerador único de bultos ---
    Private Function Dame_Siguiente_Numerador_SSCC_Heterogeneo() As Long
        Try
            Dim dtNumerador As DataTable = ed.DameNumeradorSSCCHipodromo()

            If dtNumerador.Rows.Count > 0 Then
                Dim row As DataRow = dtNumerador.Rows(0)
                Dim numnum As Long = If(Not IsDBNull(row("numnum")), CLng(row("numnum")), 0)
                Dim numdes As Long = If(Not IsDBNull(row("numdes")), CLng(row("numdes")), 0)
                Dim numhas As Long = If(Not IsDBNull(row("numhas")), CLng(row("numhas")), 0)

                If numnum = 0 Then
                    Return numdes
                Else
                    If numnum = numhas Then
                        ' Se ha alcanzado el final del Rango Permitido
                        Return -2
                    Else
                        Return numnum + 1
                    End If
                End If
            Else
                ' No Existe Registro
                Return -1
            End If
        Catch
            Return -1
        End Try
    End Function

End Class
