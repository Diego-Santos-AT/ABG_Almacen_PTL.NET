'******************************************************************************
' frmMenu.vb
'
' Form principal de la aplicación de Gestión
' Converted from VB6 to VB.NET - Line by line faithful conversion
'
' Conexiones:
'
' Creado: 8/03/00
'******************************************************************************

Imports System.Windows.Forms
Imports System.Drawing
Imports ABG_Almacen_PTL.Modules
Imports ABG_Almacen_PTL.DataAccess

Public Class frmMenu
    Inherits Form

    ' Variables igual que en VB6 (para compatibilidad)
    Private ed As EntornoDeDatos     ' Entorno de Datos de trabajo
    Private edC As edConfig          ' Entorno de configuración
    Private i As Integer             ' Variable de iteración (VB6 compatible)

    ' Botones del menú - Array como en VB6 cmdAccionMenu(0 to 5)
    Private cmdAccionMenu(5) As Button

    ' Referencia al formulario MDI padre
    Private _mdiParent As Form

    '------------------------------------------------------------------------------

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(mdiParent As Form)
        _mdiParent = mdiParent
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        ' Propiedades del formulario - Igual que VB6
        ' VB6: BackColor = &H00B06000& (BGR format) = RGB(0, 96, 176)
        Me.BackColor = Color.FromArgb(0, 96, 176)
        Me.FormBorderStyle = FormBorderStyle.None   ' BorderStyle = 0 'None
        Me.Text = "Menú Principal"                   ' Caption = "Menú Principal"
        Me.ClientSize = New Size(254, 302)          ' ClientWidth=3810, ClientHeight=4530 (twips/15)
        Me.KeyPreview = True                         ' KeyPreview = -1 'True
        Me.ShowInTaskbar = False                     ' ShowInTaskbar = 0 'False
        Me.StartPosition = FormStartPosition.Manual

        ' Configurar como hijo MDI si hay padre
        If _mdiParent IsNot Nothing Then
            Me.MdiParent = _mdiParent
        End If

        ' cmdAccionMenu(0) - CONSULTAS PTL
        ' VB6: BackColor = &H008080FF& (BGR) = RGB(255, 128, 128)
        ' Left=113, Top=75, Width=3600, Height=660 (twips/15 = pixels)
        cmdAccionMenu(0) = New Button()
        cmdAccionMenu(0).BackColor = Color.FromArgb(255, 128, 128)
        cmdAccionMenu(0).Text = "CONSULTAS PTL"
        cmdAccionMenu(0).Font = New Font("Arial", 15.75F, FontStyle.Bold)
        cmdAccionMenu(0).FlatStyle = FlatStyle.Popup
        cmdAccionMenu(0).Location = New Point(8, 5)     ' 113/15, 75/15
        cmdAccionMenu(0).Size = New Size(240, 44)        ' 3600/15, 660/15
        cmdAccionMenu(0).TabIndex = 0
        AddHandler cmdAccionMenu(0).Click, AddressOf cmdAccionMenu_Click

        ' cmdAccionMenu(1) - UBICAR BAC
        ' VB6: BackColor = &H0080C0FF& (BGR) = RGB(255, 192, 128)
        ' Left=113, Top=825
        cmdAccionMenu(1) = New Button()
        cmdAccionMenu(1).BackColor = Color.FromArgb(255, 192, 128)
        cmdAccionMenu(1).Text = "UBICAR BAC"
        cmdAccionMenu(1).Font = New Font("Arial", 15.75F, FontStyle.Bold)
        cmdAccionMenu(1).FlatStyle = FlatStyle.Popup
        cmdAccionMenu(1).Location = New Point(8, 55)    ' 113/15, 825/15
        cmdAccionMenu(1).Size = New Size(240, 44)
        cmdAccionMenu(1).TabIndex = 1
        AddHandler cmdAccionMenu(1).Click, AddressOf cmdAccionMenu_Click

        ' cmdAccionMenu(2) - EXTRAER BAC
        ' VB6: BackColor = &H0080FFFF& (BGR) = RGB(255, 255, 128)
        ' Left=113, Top=1575
        cmdAccionMenu(2) = New Button()
        cmdAccionMenu(2).BackColor = Color.FromArgb(255, 255, 128)
        cmdAccionMenu(2).Text = "EXTRAER BAC"
        cmdAccionMenu(2).Font = New Font("Arial", 15.75F, FontStyle.Bold)
        cmdAccionMenu(2).FlatStyle = FlatStyle.Popup
        cmdAccionMenu(2).Location = New Point(8, 105)   ' 113/15, 1575/15
        cmdAccionMenu(2).Size = New Size(240, 44)
        cmdAccionMenu(2).TabIndex = 2
        AddHandler cmdAccionMenu(2).Click, AddressOf cmdAccionMenu_Click

        ' cmdAccionMenu(3) - REPARTO
        ' VB6: BackColor = &H00FFFF00& (BGR) = RGB(0, 255, 255) = Yellow
        ' Left=113, Top=2325
        cmdAccionMenu(3) = New Button()
        cmdAccionMenu(3).BackColor = Color.FromArgb(0, 255, 255)
        cmdAccionMenu(3).Text = "REPARTO"
        cmdAccionMenu(3).Font = New Font("Arial", 15.75F, FontStyle.Bold)
        cmdAccionMenu(3).FlatStyle = FlatStyle.Popup
        cmdAccionMenu(3).Location = New Point(8, 155)   ' 113/15, 2325/15
        cmdAccionMenu(3).Size = New Size(240, 44)
        cmdAccionMenu(3).TabIndex = 3
        AddHandler cmdAccionMenu(3).Click, AddressOf cmdAccionMenu_Click

        ' cmdAccionMenu(4) - EMPAQUETADO
        ' VB6: BackColor = &H0080FF80& (BGR) = RGB(128, 255, 128) = Light Green
        ' Left=113, Top=3075
        cmdAccionMenu(4) = New Button()
        cmdAccionMenu(4).BackColor = Color.FromArgb(128, 255, 128)
        cmdAccionMenu(4).Text = "EMPAQUETADO"
        cmdAccionMenu(4).Font = New Font("Arial", 15.75F, FontStyle.Bold)
        cmdAccionMenu(4).FlatStyle = FlatStyle.Popup
        cmdAccionMenu(4).Location = New Point(8, 205)   ' 113/15, 3075/15
        cmdAccionMenu(4).Size = New Size(240, 44)
        cmdAccionMenu(4).TabIndex = 4
        AddHandler cmdAccionMenu(4).Click, AddressOf cmdAccionMenu_Click

        ' cmdAccionMenu(5) - SALIR
        ' VB6: No BackColor = System default (gray)
        ' Left=113, Top=3825, Height=630
        cmdAccionMenu(5) = New Button()
        cmdAccionMenu(5).BackColor = SystemColors.Control
        cmdAccionMenu(5).Text = "SALIR"
        cmdAccionMenu(5).Font = New Font("Arial", 15.75F, FontStyle.Bold)
        cmdAccionMenu(5).FlatStyle = FlatStyle.Popup
        cmdAccionMenu(5).Location = New Point(8, 255)   ' 113/15, 3825/15
        cmdAccionMenu(5).Size = New Size(240, 42)       ' 3600/15, 630/15
        cmdAccionMenu(5).TabIndex = 5
        AddHandler cmdAccionMenu(5).Click, AddressOf cmdAccionMenu_Click

        ' Agregar controles
        For Each btn In cmdAccionMenu
            Me.Controls.Add(btn)
        Next

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    '------------------------------------------------------------------------------

    Private Sub botonSalir_Click()
        Dim msg As DialogResult
        msg = MessageBox.Show("¿Desea salir de la aplicación?", "Almacén", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If msg = DialogResult.Yes Then
            If _mdiParent IsNot Nothing Then
                _mdiParent.Close()
            Else
                Application.Exit()
            End If
        End If
    End Sub

    Private Sub cmdAccionMenu_Click(sender As Object, e As EventArgs)
        Dim index As Integer = Array.IndexOf(cmdAccionMenu, sender)
        Select Case index
            Case 0
                Dim frm As New frmConsultaPTL()
                If _mdiParent IsNot Nothing Then frm.MdiParent = _mdiParent
                frm.Show()
                frm.Focus()
            Case 1
                Dim frm As New frmUbicarBAC()
                If _mdiParent IsNot Nothing Then frm.MdiParent = _mdiParent
                frm.Show()
                frm.Focus()
            Case 2
                Dim frm As New frmExtraerBAC()
                If _mdiParent IsNot Nothing Then frm.MdiParent = _mdiParent
                frm.Show()
                frm.Focus()
            Case 3
                Dim frm As New frmRepartirArticulo()
                If _mdiParent IsNot Nothing Then frm.MdiParent = _mdiParent
                frm.Show()
                frm.Focus()
            Case 4
                Dim frm As New frmEmpaquetarBAC()
                If _mdiParent IsNot Nothing Then frm.MdiParent = _mdiParent
                frm.Show()
                frm.Focus()
            Case 5
                botonSalir_Click()
        End Select
    End Sub

    Private Sub Form_Activate(sender As Object, e As EventArgs) Handles MyBase.Activated
        ' Refresca la barra de menu del formulario principal al tomar el foco
        CambiaModo(MOD_Todo)
    End Sub

    Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.D1, Keys.NumPad1
                cmdAccionMenu_Click(cmdAccionMenu(0), EventArgs.Empty)
            Case Keys.D2, Keys.NumPad2
                cmdAccionMenu_Click(cmdAccionMenu(1), EventArgs.Empty)
            Case Keys.D3, Keys.NumPad3
                cmdAccionMenu_Click(cmdAccionMenu(2), EventArgs.Empty)
            Case Keys.D4, Keys.NumPad4
                cmdAccionMenu_Click(cmdAccionMenu(3), EventArgs.Empty)
            Case Keys.D5, Keys.NumPad5
                cmdAccionMenu_Click(cmdAccionMenu(4), EventArgs.Empty)
            Case Keys.D6, Keys.NumPad6, Keys.Escape
                cmdAccionMenu_Click(cmdAccionMenu(5), EventArgs.Empty)
        End Select
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Activa el primer menu
        Me.Top = 0
        Me.Left = 0
        CargaMenu(0)

        ' ---------- Logo
        Try
            If edC IsNot Nothing Then edC.Close()
        Catch
        End Try

        edC = New edConfig()
        edC.Open()

        ' En este punto damos por buena la ejecución del programa y le damos al CargadorABG esa notificación
        ControlEjecucion()
        ' Actualización del CargadorABG
        ActualizaCargador()
    End Sub

    ' Métodos de VB6 mantenidos para compatibilidad
    Private Sub mnuArchivoSalir_Click()
        Salir()
    End Sub

    Private Sub Accion_Menu(ByVal Menu As Integer)
        Select Case Menu
            Case CMD_Salir
                Salir()
        End Select
    End Sub

    Private Sub Salir()
        Me.Close()
    End Sub

    ' Método público para compatibilidad con frmMain
    Public Sub Accion(index As Integer)
        ' Acciones de la barra de menu segun los botones pulsados
        Select Case index
            Case CMD_Salir
                Me.Close()
        End Select
    End Sub

End Class
