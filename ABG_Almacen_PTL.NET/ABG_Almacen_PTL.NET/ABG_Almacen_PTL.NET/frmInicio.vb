'*****************************************************************************************
' frmInicio.vb
' Formulario de pantalla de Inicio de la aplicación (Login)
' Muestra el nombre de programa, versión, etc
' Control de la conexión con BD
' Validación de Usuario
' Converted from VB6 to VB.NET
'*****************************************************************************************

Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports ABG_Almacen_PTL.Modules
Imports ABG_Almacen_PTL.DataAccess

Public Class frmInicio
    Inherits Form

    ' Controles del formulario
    Private WithEvents txtUsuarios As TextBox
    Private WithEvents txtPassword As TextBox
    Private WithEvents ComboEmpresa As ComboBox
    Private WithEvents ComboPuesto As ComboBox
    Private WithEvents cmdAceptar As Button
    Private WithEvents cmdCancelar As Button
    Private WithEvents Timer1 As Timer

    ' Labels
    Private lblProductName As Label
    Private lblVersion As Label
    Private lblComentarios As Label
    Private lblUsuario As Label
    Private lblContrasena As Label
    Private lblEmpresa As Label
    Private lblPuesto As Label
    Private lblEstado As Label

    ' Constantes
    Private Const CMD_Aceptar As Integer = 0
    Private Const CMD_Cancelar As Integer = 1

    ' Variables privadas
    Private Reintentos As Integer
    Private Password As String
    Private HayPassword As Boolean
    Private bEjecutado As Boolean
    Private edC As edConfig
    Private rutaLogo As String

    ' Propiedades
    Public LoginSucceeded As Boolean = False

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        ' Propiedades del formulario - Color naranja oscuro como en VB6 (&H00B06000&)
        Me.Text = "Inicio de Sesión"
        Me.FormBorderStyle = FormBorderStyle.None
        Me.BackColor = Color.FromArgb(&H0, &H60, &HB0)  ' Color azul oscuro (#0060B0)
        Me.Size = New Size(400, 350)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ShowInTaskbar = False
        Me.AutoScaleMode = AutoScaleMode.Font

        ' Título del producto
        lblProductName = New Label()
        lblProductName.Text = "ABG Almacén PTL"
        lblProductName.Font = New Font("Arial", 18, FontStyle.Bold)
        lblProductName.ForeColor = Color.White
        lblProductName.BackColor = Color.Transparent
        lblProductName.TextAlign = ContentAlignment.MiddleCenter
        lblProductName.Location = New Point(10, 20)
        lblProductName.Size = New Size(380, 35)

        ' Versión
        lblVersion = New Label()
        lblVersion.Text = "Versión 2.0.0"
        lblVersion.Font = New Font("Arial", 9)
        lblVersion.ForeColor = Color.FromArgb(&HE0, &HE0, &HE0)
        lblVersion.BackColor = Color.Transparent
        lblVersion.TextAlign = ContentAlignment.MiddleCenter
        lblVersion.Location = New Point(10, 55)
        lblVersion.Size = New Size(380, 20)

        ' Comentarios
        lblComentarios = New Label()
        lblComentarios.Text = "Sistema de Gestión de Almacén PTL"
        lblComentarios.Font = New Font("Arial", 9)
        lblComentarios.ForeColor = Color.FromArgb(&HE0, &HE0, &HE0)
        lblComentarios.BackColor = Color.Transparent
        lblComentarios.TextAlign = ContentAlignment.MiddleCenter
        lblComentarios.Location = New Point(10, 75)
        lblComentarios.Size = New Size(380, 20)

        ' Label Usuario
        lblUsuario = New Label()
        lblUsuario.Text = "Usuario:"
        lblUsuario.Font = New Font("MS Sans Serif", 10)
        lblUsuario.ForeColor = Color.FromArgb(&HE0, &HE0, &HE0)
        lblUsuario.BackColor = Color.Transparent
        lblUsuario.TextAlign = ContentAlignment.MiddleRight
        lblUsuario.Location = New Point(20, 110)
        lblUsuario.Size = New Size(100, 25)

        ' TextBox Usuario
        txtUsuarios = New TextBox()
        txtUsuarios.Font = New Font("Arial", 11, FontStyle.Bold)
        txtUsuarios.Location = New Point(130, 108)
        txtUsuarios.Size = New Size(240, 28)

        ' Label Contraseña
        lblContrasena = New Label()
        lblContrasena.Text = "Contraseña:"
        lblContrasena.Font = New Font("MS Sans Serif", 10)
        lblContrasena.ForeColor = Color.FromArgb(&HE0, &HE0, &HE0)
        lblContrasena.BackColor = Color.Transparent
        lblContrasena.TextAlign = ContentAlignment.MiddleRight
        lblContrasena.Location = New Point(20, 145)
        lblContrasena.Size = New Size(100, 25)

        ' TextBox Contraseña
        txtPassword = New TextBox()
        txtPassword.Font = New Font("Arial", 11, FontStyle.Bold)
        txtPassword.PasswordChar = "*"c
        txtPassword.Location = New Point(130, 143)
        txtPassword.Size = New Size(240, 28)

        ' Label Empresa
        lblEmpresa = New Label()
        lblEmpresa.Text = "Empresa:"
        lblEmpresa.Font = New Font("MS Sans Serif", 10)
        lblEmpresa.ForeColor = Color.FromArgb(&HE0, &HE0, &HE0)
        lblEmpresa.BackColor = Color.Transparent
        lblEmpresa.TextAlign = ContentAlignment.MiddleRight
        lblEmpresa.Location = New Point(20, 180)
        lblEmpresa.Size = New Size(100, 25)

        ' ComboBox Empresa
        ComboEmpresa = New ComboBox()
        ComboEmpresa.Font = New Font("Arial", 10)
        ComboEmpresa.DropDownStyle = ComboBoxStyle.DropDownList
        ComboEmpresa.Location = New Point(130, 178)
        ComboEmpresa.Size = New Size(240, 28)

        ' Label Puesto
        lblPuesto = New Label()
        lblPuesto.Text = "Puesto:"
        lblPuesto.Font = New Font("MS Sans Serif", 10)
        lblPuesto.ForeColor = Color.FromArgb(&HE0, &HE0, &HE0)
        lblPuesto.BackColor = Color.Transparent
        lblPuesto.TextAlign = ContentAlignment.MiddleRight
        lblPuesto.Location = New Point(20, 215)
        lblPuesto.Size = New Size(100, 25)

        ' ComboBox Puesto
        ComboPuesto = New ComboBox()
        ComboPuesto.Font = New Font("Arial", 10)
        ComboPuesto.DropDownStyle = ComboBoxStyle.DropDownList
        ComboPuesto.Location = New Point(130, 213)
        ComboPuesto.Size = New Size(240, 28)

        ' Botón Aceptar
        cmdAceptar = New Button()
        cmdAceptar.Text = "Aceptar"
        cmdAceptar.Font = New Font("Arial", 10, FontStyle.Bold)
        cmdAceptar.Location = New Point(50, 260)
        cmdAceptar.Size = New Size(130, 45)
        cmdAceptar.FlatStyle = FlatStyle.Flat
        cmdAceptar.BackColor = Color.FromArgb(80, 180, 80)
        cmdAceptar.ForeColor = Color.White

        ' Botón Cancelar
        cmdCancelar = New Button()
        cmdCancelar.Text = "Cancelar"
        cmdCancelar.Font = New Font("Arial", 10, FontStyle.Bold)
        cmdCancelar.Location = New Point(220, 260)
        cmdCancelar.Size = New Size(130, 45)
        cmdCancelar.FlatStyle = FlatStyle.Flat
        cmdCancelar.BackColor = Color.FromArgb(180, 80, 80)
        cmdCancelar.ForeColor = Color.White
        cmdCancelar.DialogResult = DialogResult.Cancel

        ' Label Estado
        lblEstado = New Label()
        lblEstado.Text = "Iniciando..."
        lblEstado.Font = New Font("Arial", 9)
        lblEstado.ForeColor = Color.FromArgb(&HC0, &HC0, &HC0)
        lblEstado.BackColor = Color.Transparent
        lblEstado.TextAlign = ContentAlignment.MiddleCenter
        lblEstado.Location = New Point(10, 315)
        lblEstado.Size = New Size(380, 20)

        ' Timer
        Timer1 = New Timer()
        Timer1.Interval = 500

        ' Agregar controles al formulario
        Me.Controls.AddRange(New Control() {
            lblProductName, lblVersion, lblComentarios,
            lblUsuario, txtUsuarios,
            lblContrasena, txtPassword,
            lblEmpresa, ComboEmpresa,
            lblPuesto, ComboPuesto,
            cmdAceptar, cmdCancelar,
            lblEstado
        })

        Me.AcceptButton = cmdAceptar
        Me.CancelButton = cmdCancelar

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Private Sub frmInicio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' FIEL A VB6: Información de versión
        Dim version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
        lblVersion.Text = $"Versión {version.Major}.{version.Minor}.{version.Build}"
        lblProductName.Text = Application.ProductName
        If String.IsNullOrEmpty(lblProductName.Text) Then
            lblProductName.Text = "ABG Almacén PTL"
        End If

        ' FIEL A VB6: RegistrarVersion (controlamos errores como en VB6)
        Try
            ControlEjecucion()
        Catch
            ' Ignorar errores como en VB6
        End Try
        
        Reintentos = 0
        LoginSucceeded = False
        bEjecutado = False

        ' Inicializar objeto de acceso a datos
        edC = New edConfig()

        ' FIEL A VB6: Mostrar mensaje de conexión
        lblEstado.Text = $"Conectando con el Servidor {BDDServLocal} ..."

        ' FIEL A VB6: Timer de inicio con 500ms
        Timer1.Interval = 500
        Timer1.Enabled = True

        txtUsuarios.Visible = True
        lblEstado.Text = "Iniciando..."
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Lanza el proceso de probar conexión solo una vez
        ' FIEL A VB6: El timer solo se ejecuta una vez
        
        If bEjecutado = False Then
            bEjecutado = True
            Timer1.Enabled = False
            
            ' Actualizar estado en la UI - FIEL A VB6
            lblEstado.Text = $"Conectando con el Servidor {BDDServLocal} ..."
            lblEstado.Refresh()
            Application.DoEvents() ' Permitir que la UI se actualice
            
            ' Probar la conexión - FIEL A VB6:
            ' En VB6, ProbarConexion devolvía False sin mensaje y luego se hacía End
            If ProbarConexion(BDDServLocal) = False Then
                ' Mostrar mensaje de error similar al comportamiento VB6
                MessageBox.Show($"No se puede conectar con el servidor {BDDServLocal}." & vbCrLf & vbCrLf &
                              "Verifique que:" & vbCrLf &
                              "- El servidor esté encendido y accesible" & vbCrLf &
                              "- El nombre del servidor en abg.ini sea correcto" & vbCrLf &
                              "- Tiene conexión a la red",
                              "Error de Conexión",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Error)
                LoginSucceeded = False
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
            Else
                lblEstado.Text = "Listo para iniciar sesión"
            End If
        End If
    End Sub

    Private Sub txtUsuarios_GotFocus(sender As Object, e As EventArgs) Handles txtUsuarios.GotFocus
        ' Usuario por defecto
        If Not String.IsNullOrEmpty(UsrDefault) Then
            txtUsuarios.Text = UsrDefault
        End If
    End Sub

    Private Sub txtUsuarios_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUsuarios.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            ValidaUsuario()
            If txtPassword.Visible Then
                txtPassword.Focus()
            Else
                cmdAceptar.Focus()
            End If
            e.Handled = True
        End If
    End Sub

    Private Sub txtUsuarios_LostFocus(sender As Object, e As EventArgs) Handles txtUsuarios.LostFocus
        ValidaUsuario()
    End Sub

    Private Sub txtPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPassword.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            cmdAceptar.Focus()
            e.Handled = True
        End If
    End Sub

    Private Sub ComboEmpresa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboEmpresa.SelectedIndexChanged
        ' FIEL A VB6: ComboEmpresa_Click
        ' En VB6 usaba On Error Resume Next para ignorar errores
        Try
            ' Validar nombre de empresa
            Dim dtEmpresa As DataTable = edC.DameCodigoEmpresa(ComboEmpresa.Text)

            If dtEmpresa.Rows.Count = 0 Then
                MessageBox.Show("Empresa desconocida. Elija una de la lista...",
                              "Conexión", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                ComboEmpresa.Focus()
                Exit Sub
            End If

            Dim empresaCod As Integer = CInt(dtEmpresa.Rows(0)("empcod"))
            Dim dtParams As DataTable = edC.DameParametrosEmpresa(empresaCod)

            If dtParams.Rows.Count > 0 Then
                If Not IsDBNull(dtParams.Rows(0)("emplog")) Then
                    rutaLogo = dtParams.Rows(0)("emplog").ToString()
                End If
            End If
        Catch ex As Exception
            ' FIEL A VB6: En VB6 usaba On Error Resume Next
            ' Ignorar errores silenciosamente
        End Try
    End Sub

    Private Sub cmdAceptar_Click(sender As Object, e As EventArgs) Handles cmdAceptar.Click
        Reintentos = Reintentos + 1
        ValidaContrasena()
    End Sub

    Private Sub cmdCancelar_Click(sender As Object, e As EventArgs) Handles cmdCancelar.Click
        ' Establecer la variable global a false para indicar un inicio de sesión fallido
        LoginSucceeded = False
        Me.Close()
    End Sub

    Private Sub ValidaUsuario()
        ' FIEL A VB6: Si el usuario está vacío, no intentar validar
        If String.IsNullOrWhiteSpace(txtUsuarios.Text) Then
            Password = ""
            If Usuario Is Nothing Then
                Usuario = New DatosUsuario()
            End If
            Usuario.Nombre = ""
            Usuario.Id = 0
            Usuario.Instancias = 0
            Usuario.NombrePC = ""
            HayPassword = False
            Me.Refresh()
            Return
        End If

        Try
            ' Comprueba si el usuario es correcto - FIEL A VB6: edC.BuscaUsuario txtUsuarios.Text
            Dim dtUsuario As DataTable = edC.BuscaUsuario(txtUsuarios.Text)

            If dtUsuario.Rows.Count > 0 Then
                Dim row As DataRow = dtUsuario.Rows(0)

                ' Obtener contraseña
                If Not IsDBNull(row("usucon")) Then
                    Password = row("usucon").ToString()
                Else
                    Password = ""
                End If

                ' Establecer datos del usuario
                If Usuario Is Nothing Then
                    Usuario = New DatosUsuario()
                End If

                Usuario.Id = CInt(row("usuide"))
                Usuario.Nombre = txtUsuarios.Text

                If Not IsDBNull(row("usuins")) Then
                    Usuario.Instancias = CInt(row("usuins"))
                Else
                    Usuario.Instancias = 1
                End If

                If Not IsDBNull(row("usunpc")) Then
                    Usuario.NombrePC = row("usunpc").ToString()
                Else
                    Usuario.NombrePC = ""
                End If

                ' Cargar empresas del usuario
                Dim empDefault As String = LeerIni(ficINI, "Varios", "EmpDefault", "")
                If String.IsNullOrEmpty(empDefault) Then
                    CargaEmpresas(0)
                Else
                    CargaEmpresas(CInt(empDefault))
                End If

                ' Cargar puestos de trabajo
                CargaPuestos(wPuestoTrabajo.Id)

                ' FIEL A VB6: Validar el nombre del PC
                ' If IsNull(Usuario.nombrePC) Or Usuario.nombrePC = "" Then ' Puede acceder a todos
                ' Else...
                If Not String.IsNullOrEmpty(Usuario.NombrePC) Then
                    Dim nombrePCActual As String = Environment.MachineName
                    If Usuario.NombrePC <> nombrePCActual Then
                        MessageBox.Show("No puede ejecutar el Programa desde este PC...",
                                      "Conexión", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Application.Exit()
                    End If
                End If

                ' Verificar si hay contraseña - FIEL A VB6
                If String.IsNullOrEmpty(Password) Then
                    HayPassword = False
                Else
                    HayPassword = True
                End If
            Else
                ' No se encuentra el usuario - FIEL A VB6
                Password = ""
                If Usuario Is Nothing Then
                    Usuario = New DatosUsuario()
                End If
                Usuario.Nombre = ""
                Usuario.Id = 0
                Usuario.Instancias = 0
                Usuario.NombrePC = ""
                HayPassword = False
            End If

            Me.Refresh()

        Catch ex As Exception
            ' FIEL A VB6: En VB6 usaba On Error Resume Next en muchos lugares
            ' No mostrar mensaje de error detallado aquí - simplemente reinicializar
            Password = ""
            If Usuario Is Nothing Then
                Usuario = New DatosUsuario()
            End If
            Usuario.Nombre = ""
            Usuario.Id = 0
            Usuario.Instancias = 0
            Usuario.NombrePC = ""
            HayPassword = False
            Me.Refresh()
        End Try
    End Sub

    Private Sub ValidaContrasena()
        Dim msg As String

        ' FIEL A VB6: Validación básica de usuario
        If String.IsNullOrWhiteSpace(txtUsuarios.Text) Then
            MessageBox.Show("Debe introducir un nombre de usuario", "Inicio de sesión",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUsuarios.Focus()
            Return
        End If

        ' FIEL A VB6: Comprobar si la contraseña es correcta
        ' If (HayPassword And txtPassword = Password) Or (Not HayPassword And Usuario.Nombre <> "") Then
        If (HayPassword AndAlso txtPassword.Text = Password) OrElse (Not HayPassword AndAlso Not String.IsNullOrEmpty(Usuario.Nombre)) Then
            LoginSucceeded = True

            lblEstado.Text = "Inicio de Sesión... "

            ' Guardar el usuario por defecto
            GuardarIni(ficINI, "Varios", "UsrDefault", txtUsuarios.Text)

            Try
                ' FIEL A VB6: Cargar parámetros de empresa
                Dim dtEmpresa As DataTable = edC.DameCodigoEmpresa(ComboEmpresa.Text)

                If dtEmpresa.Rows.Count = 0 Then
                    MessageBox.Show("Empresa desconocida. Elija una de la lista...",
                                  "Conexión", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    ComboEmpresa.Focus()
                    Exit Sub
                End If

                Dim empresaCod As Integer = CInt(dtEmpresa.Rows(0)("empcod"))
                Dim dtParams As DataTable = edC.DameParametrosEmpresa(empresaCod)

                If dtParams.Rows.Count > 0 Then
                    CodEmpresa = CInt(dtParams.Rows(0)("empcod"))
                    If Not IsDBNull(dtParams.Rows(0)("empnom")) Then
                        Empresa = dtParams.Rows(0)("empnom").ToString()
                    End If

                    ' Establecer datos de empresa
                    If EmpresaTrabajo Is Nothing Then
                        EmpresaTrabajo = New DatosEmpresaTrabajo()
                    End If
                    EmpresaTrabajo.Codigo = CodEmpresa
                    EmpresaTrabajo.Nombre = Empresa
                End If

                ' FIEL A VB6: Puesto de trabajo
                Dim dtPuesto As DataTable = edC.DameCodigoPuesto(ComboPuesto.Text)
                If dtPuesto.Rows.Count > 0 Then
                    wPuestoTrabajo.Id = CInt(dtPuesto.Rows(0)("puecod"))
                Else
                    wPuestoTrabajo.Id = 1
                End If

                ' Obtener datos del puesto de trabajo
                Dim dtPuestoInfo As DataTable = edC.DamePuestoTrabajo(wPuestoTrabajo.Id)
                If dtPuestoInfo.Rows.Count > 0 Then
                    Dim puestoRow As DataRow = dtPuestoInfo.Rows(0)
                    If Not IsDBNull(puestoRow("puedes")) Then
                        wPuestoTrabajo.Descripcion = puestoRow("puedes").ToString()
                    End If
                    If Not IsDBNull(puestoRow("puecor")) Then
                        wPuestoTrabajo.Corto = puestoRow("puecor").ToString()
                    End If
                    If Not IsDBNull(puestoRow("impcod")) Then
                        wPuestoTrabajo.Impresora = CInt(puestoRow("impcod"))
                    End If
                    If Not IsDBNull(puestoRow("impnom")) Then
                        wPuestoTrabajo.NombreImpresora = puestoRow("impnom").ToString()
                    End If
                    If Not IsDBNull(puestoRow("implen")) Then
                        wPuestoTrabajo.TipoImpresora = puestoRow("implen").ToString()
                    End If
                Else
                    wPuestoTrabajo.Descripcion = ""
                    wPuestoTrabajo.Corto = ""
                    wPuestoTrabajo.Impresora = 1
                    wPuestoTrabajo.NombreImpresora = ""
                    wPuestoTrabajo.TipoImpresora = ""
                End If

                ' Guardar puesto por defecto
                GuardarIni(ficINI, "Varios", "PueDefault", CStr(wPuestoTrabajo.Id))

                ' Impresora asociada al puesto de trabajo
                wImpresora = wPuestoTrabajo.NombreImpresora

                ' Guardar la empresa por defecto
                GuardarIni(ficINI, "Varios", "EmpDefault", CStr(CodEmpresa))

                lblEstado.Text = "Conectando con el servidor ..."
                lblEstado.Refresh()

            Catch ex As Exception
                ' FIEL A VB6: No mostrar mensaje de error, continuar
            End Try

            Me.DialogResult = DialogResult.OK
            Me.Close()
        Else
            ' FIEL A VB6: Manejo de reintentos
            If Reintentos < 3 Then
                If HayPassword Then
                    msg = "La contraseña no es válida. Vuelva a intentarlo"
                    txtPassword.Focus()
                Else
                    msg = "El nombre de usuario no es válido. Vuelva a intentarlo"
                    txtUsuarios.Focus()
                End If
                MessageBox.Show(msg, "Inicio de sesión", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                If HayPassword Then
                    msg = "La contraseña no es válida. Ha agotado los intentos de acceso"
                Else
                    msg = "El nombre de usuario no es válido. Ha agotado los intentos de acceso"
                End If
                MessageBox.Show(msg, "Inicio de sesión", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
            End If
        End If
    End Sub

    Private Sub CargaEmpresas(empdefault As Integer)
        Try
            ' Empresas de acceso del usuario
            Dim dtEmpresas As DataTable = edC.DameEmpresasAccesoUsuario(Usuario.Id)

            ' Mostrar combo
            ComboEmpresa.Visible = True
            lblEmpresa.Visible = True
            ComboEmpresa.Items.Clear()

            If dtEmpresas.Rows.Count = 0 Then
                ' No hay empresas asignadas
                MessageBox.Show("No tiene asignada empresa actualmente." & vbCrLf &
                              "Consulte con el dpto. de informática.",
                              "Conexión", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If

            Dim empdefIndex As Integer = -1
            Dim i As Integer = 0

            For Each row As DataRow In dtEmpresas.Rows
                ' Recuperar el nombre de la empresa
                Dim empresaId As Integer = CInt(row("useemp"))
                Dim dtParams As DataTable = edC.DameParametrosEmpresa(empresaId)

                If dtParams.Rows.Count > 0 Then
                    Dim nombreEmpresa As String = ""
                    If Not IsDBNull(dtParams.Rows(0)("empnom")) Then
                        nombreEmpresa = dtParams.Rows(0)("empnom").ToString()
                    End If

                    ' Verificar si es la empresa por defecto
                    If empresaId = CodEmpresa Then
                        empdefIndex = i
                        If Not IsDBNull(dtParams.Rows(0)("emplog")) Then
                            rutaLogo = dtParams.Rows(0)("emplog").ToString()
                        End If
                    End If

                    ComboEmpresa.Items.Add(nombreEmpresa)
                End If
                i += 1
            Next

            ' Seleccionar empresa por defecto
            If empdefIndex = -1 AndAlso ComboEmpresa.Items.Count > 0 Then
                ComboEmpresa.SelectedIndex = 0
            ElseIf empdefIndex >= 0 AndAlso empdefIndex < ComboEmpresa.Items.Count Then
                ComboEmpresa.SelectedIndex = empdefIndex
            End If

        Catch ex As Exception
            ' Error al cargar empresas - cargar valores por defecto
            ComboEmpresa.Items.Clear()
            ComboEmpresa.Items.Add("EMPRESA PTL")
            ComboEmpresa.SelectedIndex = 0
        End Try
    End Sub

    Private Sub CargaPuestos(idPuesto As Integer)
        Try
            Dim dtPuestos As DataTable = edC.DamePuestos()

            ' Mostrar combo
            ComboPuesto.Visible = True
            lblPuesto.Visible = True
            ComboPuesto.Items.Clear()

            If dtPuestos.Rows.Count = 0 Then
                ' No hay puestos
                ComboPuesto.Items.Add("PUESTO 1")
                ComboPuesto.SelectedIndex = 0
                Return
            End If

            Dim puestoIndex As Integer = -1
            Dim i As Integer = 0

            For Each row As DataRow In dtPuestos.Rows
                Dim puestoCod As Integer = CInt(row("puecod"))
                Dim puestoCorto As String = ""

                If Not IsDBNull(row("puecor")) Then
                    puestoCorto = row("puecor").ToString()
                End If

                ' Verificar si es el puesto por defecto
                If puestoCod = idPuesto Then
                    puestoIndex = i
                End If

                ComboPuesto.Items.Add(puestoCorto)
                i += 1
            Next

            ' Seleccionar puesto por defecto
            If puestoIndex = -1 AndAlso ComboPuesto.Items.Count > 0 Then
                ComboPuesto.SelectedIndex = 0
            ElseIf puestoIndex >= 0 AndAlso puestoIndex < ComboPuesto.Items.Count Then
                ComboPuesto.SelectedIndex = puestoIndex
            End If

        Catch ex As Exception
            ' Error al cargar puestos - cargar valores por defecto
            ComboPuesto.Items.Clear()
            ComboPuesto.Items.Add("PUESTO 1")
            ComboPuesto.Items.Add("PUESTO 2")
            ComboPuesto.Items.Add("PUESTO 3")
            ComboPuesto.SelectedIndex = 0
        End Try
    End Sub

    Private Sub frmInicio_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Timer1.Enabled = False

        ' Cerrar conexión de datos
        If edC IsNot Nothing Then
            edC.Dispose()
            edC = Nothing
        End If
    End Sub

End Class
