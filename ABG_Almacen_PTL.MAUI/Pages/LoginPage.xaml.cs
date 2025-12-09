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
                GlobalData.ConexionConfig = $"Server={GlobalData.BDDServLocal};Database={GlobalData.BDDConfig};User Id={GlobalData.UsrBDDConfig};Password={GlobalData.UsrKeyConfig};Connect Timeout={GlobalData.BDDTime};TrustServerCertificate=True;Encrypt=False;";
            }

            // Verificar conexión
            await VerificarConexion();

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
        try
        {
            lblEstado.Text = "Verificando conexión...";
            
            if (string.IsNullOrEmpty(GlobalData.ConexionConfig))
            {
                lblEstado.Text = "Error: No hay configuración de conexión";
                await DisplayAlert("Error", "No se ha configurado la conexión a la base de datos. Verifique el archivo de configuración.", "OK");
                return;
            }

            _dataAccess = new ConfigDataAccess(GlobalData.ConexionConfig);
            _dataAccess.Open();

            // Cargar puestos
            await CargarPuestos();

            lblEstado.Text = "Conectado";
            btnAceptar.IsEnabled = true;
        }
        catch (Exception ex)
        {
            lblEstado.Text = "Error de conexión";
            await DisplayAlert("Error de Conexión", 
                $"No se pudo conectar a la base de datos:\n{ex.Message}\n\nVerifique la configuración de red y base de datos.", 
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

            // Cargar empresas del usuario
            await CargarEmpresas();

            // Si hay empresa seleccionada, continuar
            if (pickerEmpresa.SelectedItem != null)
            {
                var empresaItem = (EmpresaItem)pickerEmpresa.SelectedItem;
                GlobalData.CodEmpresa = empresaItem.Id;
                ConfigurationHelper.WriteConfig("Varios", "EmpDefault", empresaItem.Id.ToString());
            }

            // Guardar puesto seleccionado
            if (pickerPuesto.SelectedItem != null)
            {
                var puestoItem = (PuestoItem)pickerPuesto.SelectedItem;
                GlobalData.wPuestoTrabajo.Id = puestoItem.Id;
                ConfigurationHelper.WriteConfig("Varios", "PueDefault", puestoItem.Id.ToString());
            }

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
