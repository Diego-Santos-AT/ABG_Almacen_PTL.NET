using System.Data;
using ABG_Almacen_PTL.MAUI.DataAccess;
using ABG_Almacen_PTL.MAUI.Modules;

namespace ABG_Almacen_PTL.MAUI.Pages;

public partial class LoginPage : ContentPage
{
    private ConfigDataAccess? _dataAccess;
    private int _reintentos = 0;
    private const int MAX_REINTENTOS = 3;
    public bool LoginSucceeded { get; private set; } = false;

    public LoginPage()
    {
        InitializeComponent();
        Initialize();
    }

    private async void Initialize()
    {
        try
        {
            // Inicializar configuración
            ConfigurationHelper.InitializeDefaultConfig();

            // Leer configuración
            GlobalData.BDDServ = ConfigurationHelper.ReadConfig("Conexion", "BDDServ", "");
            GlobalData.BDDServLocal = ConfigurationHelper.ReadConfig("Conexion", "BDDServLocal", "");

            var timeoutStr = ConfigurationHelper.ReadConfig("Conexion", "BDDTime", "30");
            if (!int.TryParse(timeoutStr, out int timeout) || timeout < 5)
            {
                timeout = 30;
            }
            GlobalData.BDDTime = timeout;

            GlobalData.BDDConfig = ConfigurationHelper.ReadConfig("Conexion", "BDDConfig", "Config");
            if (string.IsNullOrEmpty(GlobalData.BDDConfig))
            {
                GlobalData.BDDConfig = "Config";
            }

            GlobalData.UsrBDDConfig = "ABG";
            GlobalData.UsrKeyConfig = "A_34ggyx4";

            // Leer varios
            GlobalData.UsrDefault = ConfigurationHelper.ReadConfig("Varios", "UsrDefault", "");

            var empDefaultStr = ConfigurationHelper.ReadConfig("Varios", "EmpDefault", "0");
            if (!int.TryParse(empDefaultStr, out int empDefault))
            {
                empDefault = 0;
            }
            GlobalData.CodEmpresa = empDefault;

            var pueDefaultStr = ConfigurationHelper.ReadConfig("Varios", "PueDefault", "1");
            if (!int.TryParse(pueDefaultStr, out int pueDefault))
            {
                pueDefault = 1;
            }
            GlobalData.wPuestoTrabajo.Id = pueDefault;

            // Construir cadena de conexión
            if (!string.IsNullOrEmpty(GlobalData.BDDServLocal))
            {
                GlobalData.ConexionConfig = $"Server={GlobalData.BDDServLocal};Database={GlobalData.BDDConfig};User Id={GlobalData.UsrBDDConfig};Password={GlobalData.UsrKeyConfig};Connect Timeout={Constants.CTE_TimeoutPruebaConexion};TrustServerCertificate=True;Encrypt=False;";
            }

            // Verificar conexión con timeout
            await Task.Run(async () => await VerificarConexion());

            // Cargar usuario por defecto si existe
            if (!string.IsNullOrEmpty(GlobalData.UsrDefault))
            {
                txtUsuario.Text = GlobalData.UsrDefault;
                txtPassword.Focus();
            }
        }
        catch (Exception ex)
        {
            lblEstado.Text = "Error de inicialización";
            await DisplayAlert("Error", $"Error al inicializar la aplicación: {ex.Message}", "OK");
        }
    }

    private async Task VerificarConexion()
    {
        await MainThread.InvokeOnMainThreadAsync(() =>
        {
            lblEstado.Text = $"Conectando con el Servidor {GlobalData.BDDServLocal}...";
        });

        try
        {
            if (string.IsNullOrEmpty(GlobalData.ConexionConfig))
            {
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    lblEstado.Text = "Error: No hay configuración de conexión";
                });
                await DisplayAlert("Error", "No se ha configurado la conexión a la base de datos. Verifique el archivo de configuración.", "OK");
                return;
            }

            // Intentar conexión con timeout
            var connectionTask = Task.Run(() =>
            {
                _dataAccess = new ConfigDataAccess(GlobalData.ConexionConfig);
                _dataAccess.Open();
                return true;
            });

            // Esperar con timeout
            var timeoutTask = Task.Delay(TimeSpan.FromSeconds(Constants.CTE_TimeoutPruebaConexion));
            var completedTask = await Task.WhenAny(connectionTask, timeoutTask);

            if (completedTask == timeoutTask)
            {
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    lblEstado.Text = "Error de conexión: Tiempo de espera agotado";
                });
                await DisplayAlert("Error de Conexión",
                    $"No se pudo conectar al servidor {GlobalData.BDDServLocal}.\n\nVerifique que:\n- El servidor esté encendido y accesible\n- El nombre del servidor en configuración sea correcto\n- Tiene conexión a la red",
                    "OK");
                return;
            }

            if (!connectionTask.Result)
            {
                throw new Exception("No se pudo establecer la conexión");
            }

            // Cargar puestos
            await CargarPuestos();

            await MainThread.InvokeOnMainThreadAsync(() =>
            {
                lblEstado.Text = "Listo para iniciar sesión";
                btnAceptar.IsEnabled = true;
            });
        }
        catch (Exception ex)
        {
            await MainThread.InvokeOnMainThreadAsync(() =>
            {
                lblEstado.Text = "Error de conexión";
            });
            await DisplayAlert("Error de Conexión",
                $"No se pudo conectar al servidor {GlobalData.BDDServLocal}:\n{ex.Message}\n\nVerifique la configuración de red y base de datos.",
                "OK");
        }
    }

    private async Task CargarPuestos()
    {
        try
        {
            if (_dataAccess == null) return;

            var dtPuestos = _dataAccess.DamePuestos();
            if (dtPuestos.Rows.Count > 0)
            {
                var puestos = new List<PuestoItem>();
                foreach (DataRow row in dtPuestos.Rows)
                {
                    puestos.Add(new PuestoItem
                    {
                        Id = Convert.ToInt32(row["ptocod"]),
                        Descripcion = row["ptodes"].ToString() ?? ""
                    });
                }

                pickerPuesto.ItemsSource = puestos;
                pickerPuesto.ItemDisplayBinding = new Binding("Descripcion");

                // Seleccionar puesto por defecto
                if (GlobalData.wPuestoTrabajo.Id > 0)
                {
                    var puestoDefault = puestos.FirstOrDefault(p => p.Id == GlobalData.wPuestoTrabajo.Id);
                    if (puestoDefault != null)
                    {
                        pickerPuesto.SelectedItem = puestoDefault;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            await DisplayAlert("Error", $"Error al cargar puestos: {ex.Message}", "OK");
        }
    }

    private async void OnAceptarClicked(object sender, EventArgs e)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(txtUsuario.Text))
            {
                await DisplayAlert("Error", "Debe introducir un usuario", "OK");
                txtUsuario.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                await DisplayAlert("Error", "Debe introducir una contraseña", "OK");
                txtPassword.Focus();
                return;
            }

            btnAceptar.IsEnabled = false;
            lblEstado.Text = "Validando usuario...";

            if (_dataAccess == null)
            {
                await DisplayAlert("Error", "No hay conexión a la base de datos", "OK");
                btnAceptar.IsEnabled = true;
                return;
            }

            // Buscar usuario
            var dtUsuario = _dataAccess.BuscaUsuario(txtUsuario.Text);
            
            if (dtUsuario.Rows.Count == 0)
            {
                _reintentos++;
                if (_reintentos >= MAX_REINTENTOS)
                {
                    await DisplayAlert("Error", "Número máximo de intentos alcanzado", "OK");
                    Application.Current?.Quit();
                    return;
                }

                await DisplayAlert("Error", "Usuario no encontrado", "OK");
                txtUsuario.Focus();
                btnAceptar.IsEnabled = true;
                return;
            }

            var row = dtUsuario.Rows[0];
            var passwordDB = row["usucla"]?.ToString() ?? "";

            if (passwordDB != txtPassword.Text)
            {
                _reintentos++;
                if (_reintentos >= MAX_REINTENTOS)
                {
                    await DisplayAlert("Error", "Número máximo de intentos alcanzado", "OK");
                    Application.Current?.Quit();
                    return;
                }

                await DisplayAlert("Error", "Contraseña incorrecta", "OK");
                txtPassword.Text = "";
                txtPassword.Focus();
                btnAceptar.IsEnabled = true;
                return;
            }

            // Login exitoso
            GlobalData.Usuario.Id = Convert.ToInt32(row["usucod"]);
            GlobalData.Usuario.Nombre = row["usunom"]?.ToString() ?? "";
            GlobalData.Usuario.Clave = passwordDB;

            // Guardar usuario por defecto
            ConfigurationHelper.WriteConfig("Varios", "UsrDefault", txtUsuario.Text);

            // Cargar empresas del usuario - IMPORTANTE: Debe hacerse después de validar el usuario
            lblEstado.Text = "Cargando empresas...";
            await CargarEmpresas();

            // Validar que se haya seleccionado una empresa
            if (pickerEmpresa.SelectedItem == null)
            {
                if (pickerEmpresa.ItemsSource is List<EmpresaItem> empresas && empresas.Count == 0)
                {
                    await DisplayAlert("Error", "No tiene asignada ninguna empresa.\nConsulte con el departamento de informática.", "OK");
                    btnAceptar.IsEnabled = true;
                    return;
                }
                else
                {
                    await DisplayAlert("Error", "Debe seleccionar una empresa", "OK");
                    pickerEmpresa.Focus();
                    btnAceptar.IsEnabled = true;
                    return;
                }
            }

            // Validar que se haya seleccionado un puesto
            if (pickerPuesto.SelectedItem == null)
            {
                await DisplayAlert("Error", "Debe seleccionar un puesto de trabajo", "OK");
                pickerPuesto.Focus();
                btnAceptar.IsEnabled = true;
                return;
            }

            // Guardar empresa seleccionada
            var empresaItem = (EmpresaItem)pickerEmpresa.SelectedItem;
            GlobalData.CodEmpresa = empresaItem.Id;
            GlobalData.Empresa = empresaItem.Nombre;
            ConfigurationHelper.WriteConfig("Varios", "EmpDefault", empresaItem.Id.ToString());

            // Guardar puesto seleccionado
            var puestoItem = (PuestoItem)pickerPuesto.SelectedItem;
            GlobalData.wPuestoTrabajo.Id = puestoItem.Id;
            GlobalData.wPuestoTrabajo.Descripcion = puestoItem.Descripcion;
            ConfigurationHelper.WriteConfig("Varios", "PueDefault", puestoItem.Id.ToString());

            LoginSucceeded = true;
            GlobalData.LoginSucceeded = true;

            lblEstado.Text = "Login exitoso";

            // Navegar a la página principal
            await Shell.Current.GoToAsync("//MainPage");
        }
        catch (Exception ex)
        {
            lblEstado.Text = "Error de login";
            await DisplayAlert("Error", $"Error durante el login: {ex.Message}", "OK");
            btnAceptar.IsEnabled = true;
        }
    }

    private async Task CargarEmpresas()
    {
        try
        {
            if (_dataAccess == null) return;

            var dtEmpresas = _dataAccess.DameEmpresasAccesoUsuario(GlobalData.Usuario.Id);
            if (dtEmpresas.Rows.Count > 0)
            {
                var empresas = new List<EmpresaItem>();
                foreach (DataRow row in dtEmpresas.Rows)
                {
                    empresas.Add(new EmpresaItem
                    {
                        Id = Convert.ToInt32(row["empcod"]),
                        Nombre = row["empnom"]?.ToString() ?? ""
                    });
                }

                pickerEmpresa.ItemsSource = empresas;
                pickerEmpresa.ItemDisplayBinding = new Binding("Nombre");

                // Seleccionar empresa por defecto
                if (GlobalData.CodEmpresa > 0)
                {
                    var empresaDefault = empresas.FirstOrDefault(e => e.Id == GlobalData.CodEmpresa);
                    if (empresaDefault != null)
                    {
                        pickerEmpresa.SelectedItem = empresaDefault;
                    }
                }
                else if (empresas.Count == 1)
                {
                    pickerEmpresa.SelectedItem = empresas[0];
                }
            }
        }
        catch (Exception ex)
        {
            await DisplayAlert("Error", $"Error al cargar empresas: {ex.Message}", "OK");
        }
    }

    private void OnCancelarClicked(object sender, EventArgs e)
    {
        Application.Current?.Quit();
    }

    private void OnUsuarioCompleted(object sender, EventArgs e)
    {
        txtPassword.Focus();
    }

    private void OnPasswordCompleted(object sender, EventArgs e)
    {
        if (btnAceptar.IsEnabled)
        {
            OnAceptarClicked(sender, e);
        }
    }

    // Clases auxiliares para los Pickers
    private class EmpresaItem
    {
        public int Id { get; set; }
        public string Nombre { get; set; } = "";
    }

    private class PuestoItem
    {
        public int Id { get; set; }
        public string Descripcion { get; set; } = "";
    }
}
