using ABG_Almacen_PTL.MAUI.DataAccess;
using ABG_Almacen_PTL.MAUI.Modules;
using ABG_Almacen_PTL.MAUI.Models;
using System.Collections.ObjectModel;
using System.Data;

namespace ABG_Almacen_PTL.MAUI.Pages;

public partial class EmpaquetarBACPage : ContentPage
{
    private PTLDataAccess? _dataAccess;
    private ObservableCollection<ArticuloItem> _articulos = new();
    private string _bacActual = "";
    private int _estadoBAC = 0;
    private int _ubicacionBAC = 0;

    public EmpaquetarBACPage()
    {
        InitializeComponent();
        collectionArticulos.ItemsSource = _articulos;
    }

    protected override void OnAppearing()
    {
        base.OnAppearing();
        try
        {
            _dataAccess = new PTLDataAccess();
            _dataAccess.Open();
            LimpiarDatos();
            txtLecturaCodigo.Focus();
        }
        catch (Exception ex)
        {
            DisplayAlert("Error", $"Error de conexión: {ex.Message}", "OK");
        }
    }

    protected override void OnDisappearing()
    {
        base.OnDisappearing();
        _dataAccess?.Dispose();
        _dataAccess = null;
    }

    private void OnCodigoCompleted(object? sender, EventArgs e)
    {
        string codigo = txtLecturaCodigo.Text?.Trim() ?? "";
        
        if (string.IsNullOrEmpty(codigo))
            return;

        LimpiarDatos();

        try
        {
            switch (codigo.Length)
            {
                case 12: // BAC
                    lblTipoLabel.Text = "BAC:";
                    ValidarBAC(codigo);
                    break;

                case 18: // SSCC de Caja
                    lblTipoLabel.Text = "CAJA:";
                    ValidarCaja(codigo);
                    break;

                case 20: // SSCC de Caja con prefijo
                    lblTipoLabel.Text = "CAJA:";
                    ValidarCaja(codigo.Substring(2));
                    break;
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }

        txtLecturaCodigo.Text = "";
    }

    private void ValidarBAC(string bac)
    {
        if (_dataAccess == null) return;

        try
        {
            var dtBAC = _dataAccess.DameDatosBACdePTL(bac);

            if (dtBAC.Rows.Count > 0)
            {
                var row = dtBAC.Rows[0];

                double unipes = row["unipes"] != DBNull.Value ? Convert.ToDouble(row["unipes"]) : 0;
                double univol = row["univol"] != DBNull.Value ? Convert.ToDouble(row["univol"]) : 0;
                string unicod = row["unicod"] != DBNull.Value ? row["unicod"].ToString() ?? "" : "";
                _estadoBAC = row["uniest"] != DBNull.Value ? Convert.ToInt32(row["uniest"]) : 0;
                _ubicacionBAC = row["uninum"] != DBNull.Value ? Convert.ToInt32(row["uninum"]) : 0;
                int unigru = row["unigru"] != DBNull.Value ? Convert.ToInt32(row["unigru"]) : 0;
                int unitab = row["unitab"] != DBNull.Value ? Convert.ToInt32(row["unitab"]) : 0;
                string unicaj = row["unicaj"] != DBNull.Value ? row["unicaj"].ToString() ?? "" : "";

                _bacActual = unicod;

                MainThread.BeginInvokeOnMainThread(() =>
                {
                    lblCodigo.Text = unicod;
                    lblGrupo.Text = unigru.ToString();
                    lblTablilla.Text = unitab.ToString();
                    lblNumCaja.Text = unicaj;
                    lblPeso.Text = unipes.ToString("F3");
                    lblVolumen.Text = univol.ToString("F3");
                    btnEmpaquetar.IsEnabled = true;
                });

                // Cargar artículos del BAC
                CargarArticulosBAC(unigru, unicod);

                // Comprobar estado
                if (_estadoBAC == 0 && chkCerrarBAC.IsChecked)
                {
                    CerrarBAC();
                }

                if (_ubicacionBAC > 0 && chkExtraerBAC.IsChecked)
                {
                    ExtraerBAC();
                }
            }
            else
            {
                MostrarMensaje("No existe el BAC", TipoMensaje.MENSAJE_Grave);
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }
    }

    private void ValidarCaja(string sscc)
    {
        if (_dataAccess == null) return;

        try
        {
            var dtCaja = _dataAccess.DameDatosCAJAdePTL(sscc);

            if (dtCaja.Rows.Count > 0)
            {
                var row = dtCaja.Rows[0];
                
                int ltcgru = row["ltcgru"] != DBNull.Value ? Convert.ToInt32(row["ltcgru"]) : 0;
                int ltctab = row["ltctab"] != DBNull.Value ? Convert.ToInt32(row["ltctab"]) : 0;
                string ltccaj = row["ltccaj"] != DBNull.Value ? row["ltccaj"].ToString() ?? "" : "";
                double ltcpes = row["ltcpes"] != DBNull.Value ? Convert.ToDouble(row["ltcpes"]) : 0;
                double ltcvol = row["ltcvol"] != DBNull.Value ? Convert.ToDouble(row["ltcvol"]) : 0;

                MainThread.BeginInvokeOnMainThread(() =>
                {
                    lblCodigo.Text = sscc;
                    lblGrupo.Text = ltcgru.ToString();
                    lblTablilla.Text = ltctab.ToString();
                    lblNumCaja.Text = ltccaj;
                    lblPeso.Text = ltcpes.ToString("F3");
                    lblVolumen.Text = ltcvol.ToString("F3");
                });

                // Cargar artículos de la caja
                CargarArticulosCaja(ltcgru, ltctab, ltccaj);
            }
            else
            {
                MostrarMensaje("No existe la CAJA", TipoMensaje.MENSAJE_Grave);
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }
    }

    private void CargarArticulosBAC(int grupo, string bac)
    {
        if (_dataAccess == null) return;

        try
        {
            var dtContenido = _dataAccess.DameContenidoBacGrupo(grupo, bac);
            
            MainThread.BeginInvokeOnMainThread(() =>
            {
                _articulos.Clear();
                int totalUds = 0;

                foreach (DataRow row in dtContenido.Rows)
                {
                    string codigo = row["uniart"] != DBNull.Value ? row["uniart"].ToString() ?? "" : "";
                    string nombre = row["artnom"] != DBNull.Value ? row["artnom"].ToString() ?? "" : "";
                    int cantidad = row["unican"] != DBNull.Value ? Convert.ToInt32(row["unican"]) : 0;

                    _articulos.Add(new ArticuloItem
                    {
                        Codigo = codigo,
                        Nombre = nombre,
                        Cantidad = cantidad.ToString()
                    });

                    totalUds += cantidad;
                }

                lblUds.Text = totalUds.ToString();
                lblArts.Text = _articulos.Count.ToString();
            });
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error al cargar artículos: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }
    }

    private void CargarArticulosCaja(int grupo, int tablilla, string caja)
    {
        if (_dataAccess == null) return;

        try
        {
            var dtContenido = _dataAccess.DameContenidoCajaGrupo(grupo, tablilla, caja);
            
            MainThread.BeginInvokeOnMainThread(() =>
            {
                _articulos.Clear();
                int totalUds = 0;

                foreach (DataRow row in dtContenido.Rows)
                {
                    string codigo = row["ltaart"] != DBNull.Value ? row["ltaart"].ToString() ?? "" : "";
                    string nombre = row["artnom"] != DBNull.Value ? row["artnom"].ToString() ?? "" : "";
                    double cantidad = row["ltacan"] != DBNull.Value ? Convert.ToDouble(row["ltacan"]) : 0;

                    _articulos.Add(new ArticuloItem
                    {
                        Codigo = codigo,
                        Nombre = nombre,
                        Cantidad = cantidad.ToString()
                    });

                    totalUds += (int)cantidad;
                }

                lblUds.Text = totalUds.ToString();
                lblArts.Text = _articulos.Count.ToString();
            });
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error al cargar artículos: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }
    }

    private void CerrarBAC()
    {
        if (_dataAccess == null || string.IsNullOrEmpty(_bacActual)) return;

        try
        {
            _dataAccess.CambiaEstadoBACdePTL(_bacActual, 1, GlobalData.Usuario.Id, out int retorno, out string msgSalida);
            
            if (retorno == 0)
            {
                _estadoBAC = 1;
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error al cerrar BAC: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }
    }

    private void ExtraerBAC()
    {
        if (_dataAccess == null || string.IsNullOrEmpty(_bacActual)) return;

        try
        {
            _dataAccess.RetirarBACdePTL(_bacActual, GlobalData.Usuario.Id, out int retorno, out string msgSalida);
            
            if (retorno == 0)
            {
                _ubicacionBAC = 0;
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error al extraer BAC: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }
    }

    private async void OnEmpaquetarClicked(object? sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(_bacActual) || _dataAccess == null)
        {
            MostrarMensaje("No hay BAC seleccionado", TipoMensaje.MENSAJE_Grave);
            return;
        }

        if (_articulos.Count == 0)
        {
            MostrarMensaje("El BAC no tiene artículos para empaquetar", TipoMensaje.MENSAJE_Grave);
            return;
        }

        try
        {
            // Generar SSCC (simplificado)
            string sscc = ObtenerSSCC();
            
            if (string.IsNullOrEmpty(sscc))
            {
                MostrarMensaje("No se pudo generar SSCC para la caja", TipoMensaje.MENSAJE_Grave);
                return;
            }

            // Empaquetar BAC a CAJA
            _dataAccess.TraspasaBACaCAJAdePTLByRef(_bacActual, GlobalData.Usuario.Id, sscc, out int retorno, out string msgSalida);

            if (retorno == 0)
            {
                MostrarMensaje($"Se ha empaquetado el BAC exitosamente.\nSSCC: {sscc}", TipoMensaje.MENSAJE_Exclamacion);
                
                // Mostrar datos de la caja creada
                lblTipoLabel.Text = "CAJA:";
                ValidarCaja(sscc);

                if (chkImprimirEtiqueta.IsChecked)
                {
                    await DisplayAlert("Imprimir", "Etiqueta enviada a imprimir", "OK");
                }
            }
            else
            {
                MostrarMensaje($"Error al empaquetar BAC: {msgSalida}", TipoMensaje.MENSAJE_Grave);
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }
    }

    private string ObtenerSSCC()
    {
        if (_dataAccess == null) return "";

        try
        {
            var dtNumerador = _dataAccess.DameNumeradorSSCCHipodromo();
            
            if (dtNumerador.Rows.Count > 0)
            {
                var row = dtNumerador.Rows[0];
                long numnum = row["numnum"] != DBNull.Value ? Convert.ToInt64(row["numnum"]) : 0;
                long numdes = row["numdes"] != DBNull.Value ? Convert.ToInt64(row["numdes"]) : 0;
                long numhas = row["numhas"] != DBNull.Value ? Convert.ToInt64(row["numhas"]) : 0;

                long siguiente = numnum == 0 ? numdes : numnum + 1;

                if (siguiente > numhas)
                {
                    return "";
                }

                // Actualizar numerador
                _dataAccess.ActualizaNumeradorSSCCHipodromo((int)siguiente);

                // Generar SSCC
                return "3842" + siguiente.ToString().PadLeft(14, '0');
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error al obtener SSCC: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }

        return "";
    }

    private void LimpiarDatos()
    {
        _bacActual = "";
        _estadoBAC = 0;
        _ubicacionBAC = 0;

        MainThread.BeginInvokeOnMainThread(() =>
        {
            lblTipoLabel.Text = "BAC:";
            lblCodigo.Text = "";
            lblGrupo.Text = "";
            lblTablilla.Text = "";
            lblNumCaja.Text = "";
            lblUds.Text = "";
            lblArts.Text = "";
            lblPeso.Text = "";
            lblVolumen.Text = "";
            btnEmpaquetar.IsEnabled = false;
            _articulos.Clear();
        });
    }

    private void MostrarMensaje(string mensaje, TipoMensaje tipo)
    {
        string titulo = tipo switch
        {
            TipoMensaje.MENSAJE_Informativo => "Información",
            TipoMensaje.MENSAJE_Grave => "Error",
            TipoMensaje.MENSAJE_Exclamacion => "Aviso",
            _ => "Mensaje"
        };

        MainThread.BeginInvokeOnMainThread(async () =>
        {
            await DisplayAlert(titulo, mensaje, "OK");
        });
    }

    private async void OnSalirClicked(object? sender, EventArgs e)
    {
        await Shell.Current.GoToAsync("..");
    }

    private void OnEntryFocused(object? sender, FocusEventArgs e)
    {
        if (sender is Entry entry)
        {
            entry.BackgroundColor = Colors.LightGreen;
        }
    }

    private void OnEntryUnfocused(object? sender, FocusEventArgs e)
    {
        if (sender is Entry entry)
        {
            entry.BackgroundColor = Colors.White;
        }
    }
}
