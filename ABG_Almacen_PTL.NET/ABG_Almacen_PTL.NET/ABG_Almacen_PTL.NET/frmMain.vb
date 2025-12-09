'******************************************************************************
' frmMain.vb
'
' Form principal de la aplicación de Gestión de Almacén PTL
' Converted from VB6 to VB.NET - Faithful line-by-line conversion
'
' Creado   : 30/01/2001
' Ult. Mod.: 23/09/2020
'******************************************************************************

Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing
Imports ABG_Almacen_PTL.Modules
Imports ABG_Almacen_PTL.DataAccess

Public Class frmMain
    Inherits Form

    ' Declaraciones API igual que en VB6
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    End Function

    Private Const EM_UNDO As Integer = &HC7

    ' Variables igual que en VB6
    Private edC As edConfig

    ' Controles principales
    Private WithEvents tbToolBar As ToolStrip
    Private WithEvents sbStatusBar As StatusStrip
    Private statusPanels(4) As ToolStripStatusLabel

    ' Formulario de menú
    Private _frmMenu As frmMenu

    '------------------------------------------------------------------------------

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        ' Propiedades del formulario MDI - igual que VB6
        ' VB6: BackColor = &H8000000C& (System AppWorkspace)
        Me.BackColor = SystemColors.AppWorkspace
        Me.Text = "PTL ALM"
        Me.IsMdiContainer = True
        Me.WindowState = FormWindowState.Maximized  ' VB6: WindowState = 2 'Maximized
        Me.StartPosition = FormStartPosition.CenterScreen  ' VB6: StartUpPosition = 1

        ' Crear ToolBar (tbToolbar) - Visible = False en VB6
        tbToolBar = New ToolStrip()
        tbToolBar.Dock = DockStyle.Top
        tbToolBar.Visible = False  ' VB6: Visible = 0 'False

        ' Crear StatusBar (sbStatusBar) - Visible = False en VB6
        sbStatusBar = New StatusStrip()
        sbStatusBar.Dock = DockStyle.Bottom
        sbStatusBar.Visible = False  ' VB6: Visible = 0 'False

        ' VB6 tiene 5 paneles en sbStatusBar
        For i As Integer = 0 To 4
            statusPanels(i) = New ToolStripStatusLabel()
            statusPanels(i).BorderSides = ToolStripStatusLabelBorderSides.All
            sbStatusBar.Items.Add(statusPanels(i))
        Next

        ' Panel 1: AutoSize = 1 (Spring)
        statusPanels(0).Spring = True

        ' Panel 5: Style = 6 (Date), Alignment = 1
        statusPanels(4).Text = Date.Today.ToString("dd/MM/yyyy")

        ' Agregar controles al formulario
        Me.Controls.Add(tbToolBar)
        Me.Controls.Add(sbStatusBar)

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    '------------------------------------------------------------------------------

    Private Sub MDIForm_Click(sender As Object, e As EventArgs) Handles Me.Click
        LoadMenu()
    End Sub

    Private Sub MDIForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Carga del ini la pantalla
        Dim mainLeft As Integer = CInt(LeerIni(ficINI, "Pantalla", "MainLeft", "1000"))
        Dim mainTop As Integer = CInt(LeerIni(ficINI, "Pantalla", "MainTop", "1000"))
        Dim mainWidth As Integer = CInt(LeerIni(ficINI, "Pantalla", "MainWidth", "3735"))
        Dim mainHeight As Integer = CInt(LeerIni(ficINI, "Pantalla", "MainHeight", "4860"))

        ' Convertir de twips a pixels (aprox /15)
        Me.Left = mainLeft \ 15
        Me.Top = mainTop \ 15
        Me.Width = mainWidth \ 15
        Me.Height = mainHeight \ 15

        Me.Text = "PTL ALM (" & Empresa & ")"
        statusPanels(EST_Empresa - 1).Text = Empresa
        statusPanels(EST_Usuario - 1).Text = Usuario.Nombre

        ' -- Carga el formulario menú
        LoadMenu()
    End Sub

    Private Sub MDIForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' VB6: MDIForm_QueryUnload
        If e.CloseReason <> CloseReason.UserClosing Then Exit Sub

        If MessageBox.Show("¿Desea salir de la aplicación?", "Almacén",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            e.Cancel = True
        End If
    End Sub

    Private Sub MDIForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ' VB6: MDIForm_Unload
        If Me.WindowState <> FormWindowState.Minimized Then
            ' Guardar posición en twips (multiplicar por 15)
            GuardarIni(ficINI, "Pantalla", "MainLeft", CStr(Me.Left * 15))
            GuardarIni(ficINI, "Pantalla", "MainTop", CStr(Me.Top * 15))
            GuardarIni(ficINI, "Pantalla", "MainWidth", CStr(Me.Width * 15))
            GuardarIni(ficINI, "Pantalla", "MainHeight", CStr(Me.Height * 15))
        End If
    End Sub

    Private Sub LoadMenu()
        If _frmMenu Is Nothing OrElse _frmMenu.IsDisposed Then
            _frmMenu = New frmMenu(Me)
        End If
        _frmMenu.Show()
    End Sub

    Private Sub tbToolBar_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles tbToolBar.ItemClicked
        Dim index As Integer = tbToolBar.Items.IndexOf(e.ClickedItem)
        Accion(index + 1)  ' VB6 es 1-based
    End Sub

    Public Sub Accion(index As Integer)
        Try
            ' Acciones de la barra de menu segun los botones pulsados
            Select Case index
                Case CMD_Menu
                    Dim activeChild As Form = Me.ActiveMdiChild
                    If activeChild Is Nothing OrElse activeChild.Text <> "Menú Principal" Then
                        LoadMenu()
                        _frmMenu.Focus()
                    End If
                Case Else
                    Dim activeChild As Form = Me.ActiveMdiChild
                    If index = CMD_Salir AndAlso (activeChild Is Nothing OrElse activeChild.Text = Me.Text) Then
                        Me.Close()
                    Else
                        ' Llamar Accion del formulario hijo activo si tiene el método
                        If activeChild IsNot Nothing Then
                            Dim accionMethod = activeChild.GetType().GetMethod("Accion")
                            If accionMethod IsNot Nothing Then
                                accionMethod.Invoke(activeChild, New Object() {index})
                            End If
                        End If
                    End If
            End Select
        Catch
            ' On Error Resume Next en VB6
        End Try
    End Sub

    ' Método público para cambiar el modo del menú (compatible con VB6)
    Public Sub CambiaModoMenu(Modo As Integer)
        CambiaModo(Modo, tbToolBar)
    End Sub

End Class
