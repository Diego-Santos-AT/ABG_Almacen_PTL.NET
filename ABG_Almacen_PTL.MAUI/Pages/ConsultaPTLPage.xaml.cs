using ABG_Almacen_PTL.MAUI.DataAccess;
using ABG_Almacen_PTL.MAUI.Modules;
using ABG_Almacen_PTL.MAUI.Models;
using System.Collections.ObjectModel;
using System.Data;

namespace ABG_Almacen_PTL.MAUI.Pages;

public partial class ConsultaPTLPage : ContentPage
{
    private PTLDataAccess? _dataAccess;
    private ObservableCollection<ArticuloItem> _articulos = new();

    public ConsultaPTLPage()
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
                case 12: // Unidad de transporte / Ubicación
                    lblTipo.Text = "BAC";
                    if (!ValidarBAC(codigo))
                    {
                        if (!ValidarUbicacion(codigo))
                        {
                            MostrarMensaje("No se ha encontrado Ubicación o BAC", TipoMensaje.MENSAJE_Grave);
                        }
                    }
                    break;

                case 18: // SSCC de Caja
                    lblTipo.Text = "CAJA";
                    ValidarCaja(codigo);
                    break;

                case 20: // SSCC de Caja con prefijo
                    lblTipo.Text = "CAJA";
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

    private bool ValidarBAC(string bac)
    {
        if (_dataAccess == null) return false;

        try
        {
            var dtBAC = _dataAccess.DameDatosBACdePTL(bac);

            if (dtBAC.Rows.Count > 0)
            {
                var row = dtBAC.Rows[0];
                MostrarDatosBAC(row);
                return true;
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }

        return false;
    }

    private bool ValidarUbicacion(string ubicacion)
    {
        if (_dataAccess == null) return false;

        try
        {
            int iALF = 2;
            int iALM = int.Parse(ubicacion.Substring(0, 3));
            int iBLO = int.Parse(ubicacion.Substring(3, 3));
            int iFIL = int.Parse(ubicacion.Substring(6, 3));
            int iALT = int.Parse(ubicacion.Substring(9, 3));

            var dtUbicacion = _dataAccess.DameDatosUbicacionPTL(iALF, iALM, iBLO, iFIL, iALT);

            if (dtUbicacion.Rows.Count > 0)
            {
                var row = dtUbicacion.Rows[0];
                
                int ubicod = row["ubicod"] != DBNull.Value ? Convert.ToInt32(row["ubicod"]) : 0;
                MainThread.BeginInvokeOnMainThread(() =>
                {
                    lblUbicacion.Text = $"({ubicod}) {iALM:000}.{iBLO:000}.{iFIL:000}.{iALT:000}";
                });

                if (row["unicod"] != DBNull.Value)
                {
                    string unicod = row["unicod"].ToString() ?? "";
                    ValidarBAC(unicod);
                }

                return true;
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }

        return false;
    }

    private bool ValidarCaja(string sscc)
    {
        if (_dataAccess == null) return false;

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

                return true;
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

        return false;
    }

    private void MostrarDatosBAC(DataRow row)
    {
        double unipes = row["unipes"] != DBNull.Value ? Convert.ToDouble(row["unipes"]) : 0;
        double univol = row["univol"] != DBNull.Value ? Convert.ToDouble(row["univol"]) : 0;
        string unicod = row["unicod"] != DBNull.Value ? row["unicod"].ToString() ?? "" : "";
        int unigru = row["unigru"] != DBNull.Value ? Convert.ToInt32(row["unigru"]) : 0;
        int unitab = row["unitab"] != DBNull.Value ? Convert.ToInt32(row["unitab"]) : 0;
        string unicaj = row["unicaj"] != DBNull.Value ? row["unicaj"].ToString() ?? "" : "";

        MainThread.BeginInvokeOnMainThread(() =>
        {
            if (row["ubicod"] != DBNull.Value)
            {
                int ubicod = Convert.ToInt32(row["ubicod"]);
                int ubialm = row["ubialm"] != DBNull.Value ? Convert.ToInt32(row["ubialm"]) : 0;
                int ubiblo = row["ubiblo"] != DBNull.Value ? Convert.ToInt32(row["ubiblo"]) : 0;
                int ubifil = row["ubifil"] != DBNull.Value ? Convert.ToInt32(row["ubifil"]) : 0;
                int ubialt = row["ubialt"] != DBNull.Value ? Convert.ToInt32(row["ubialt"]) : 0;
                lblUbicacion.Text = $"({ubicod}) {ubialm:000}.{ubiblo:000}.{ubifil:000}.{ubialt:000}";
            }
            else
            {
                lblUbicacion.Text = "SIN UBICACION";
            }

            lblCodigo.Text = unicod;
            lblGrupo.Text = unigru.ToString();
            lblTablilla.Text = unitab.ToString();
            lblNumCaja.Text = unicaj;
            lblPeso.Text = unipes.ToString("F3");
            lblVolumen.Text = univol.ToString("F3");
        });

        // Cargar artículos del BAC
        CargarArticulosBAC(unigru, unicod);
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

    private void LimpiarDatos()
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            lblTipo.Text = "BAC";
            lblUbicacion.Text = "";
            lblCodigo.Text = "";
            lblGrupo.Text = "";
            lblTablilla.Text = "";
            lblNumCaja.Text = "";
            lblUds.Text = "";
            lblArts.Text = "";
            lblPeso.Text = "";
            lblVolumen.Text = "";
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
