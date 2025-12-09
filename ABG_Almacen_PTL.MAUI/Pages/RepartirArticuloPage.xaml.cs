using ABG_Almacen_PTL.MAUI.DataAccess;
using ABG_Almacen_PTL.MAUI.Modules;

namespace ABG_Almacen_PTL.MAUI.Pages;

public partial class RepartirArticuloPage : ContentPage
{
    private PTLDataAccess? _dataAccess;
    private string _articuloActual = "";
    private string _bacDestino = "";
    private int _grupoBAC = 0;
    private int _tablillaBAC = 0;

    public RepartirArticuloPage()
    {
        InitializeComponent();
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

        // TODO: Implement full article lookup from database
        // This requires additional stored procedures not present in the VB.NET version
        // For now, we display the scanned code as a placeholder
        
        MainThread.BeginInvokeOnMainThread(() =>
        {
            _articuloActual = codigo;
            lblArticulo.Text = codigo;
            lblNombreArticulo.Text = "Artículo escaneado - Implementar búsqueda completa";
            lblEAN13.Text = codigo.Length == 13 ? codigo : "N/A";
            lblPeso.Text = "0.000";
            lblVolumen.Text = "0.000";
        });

        txtLecturaCodigo.Text = "";
        txtBAC.Focus();
    }

    private void OnBACCompleted(object? sender, EventArgs e)
    {
        string bac = txtBAC.Text?.Trim() ?? "";
        
        if (string.IsNullOrEmpty(bac))
            return;

        if (_dataAccess == null) return;

        try
        {
            var dtBAC = _dataAccess.DameDatosBACdePTL(bac);

            if (dtBAC.Rows.Count > 0)
            {
                var row = dtBAC.Rows[0];
                
                string unicod = row["unicod"] != DBNull.Value ? row["unicod"].ToString() ?? "" : "";
                int unigru = row["unigru"] != DBNull.Value ? Convert.ToInt32(row["unigru"]) : 0;
                int unitab = row["unitab"] != DBNull.Value ? Convert.ToInt32(row["unitab"]) : 0;

                _bacDestino = unicod;
                _grupoBAC = unigru;
                _tablillaBAC = unitab;

                MainThread.BeginInvokeOnMainThread(() =>
                {
                    lblBACDestino.Text = unicod;
                    lblGrupoBAC.Text = unigru.ToString();
                    lblTablillaBAC.Text = unitab.ToString();
                });

                txtCantidad.Focus();
            }
            else
            {
                MostrarMensaje("No existe el BAC", TipoMensaje.MENSAJE_Grave);
                txtBAC.Text = "";
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }
    }

    private void OnRestarClicked(object? sender, EventArgs e)
    {
        if (int.TryParse(txtCantidad.Text, out int cantidad))
        {
            if (cantidad > 1)
            {
                txtCantidad.Text = (cantidad - 1).ToString();
            }
        }
    }

    private void OnSumarClicked(object? sender, EventArgs e)
    {
        if (int.TryParse(txtCantidad.Text, out int cantidad))
        {
            txtCantidad.Text = (cantidad + 1).ToString();
        }
        else
        {
            txtCantidad.Text = "1";
        }
    }

    private void OnRepartirClicked(object? sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(_articuloActual))
        {
            MostrarMensaje("Debe escanear un artículo primero", TipoMensaje.MENSAJE_Grave);
            return;
        }

        if (string.IsNullOrEmpty(_bacDestino))
        {
            MostrarMensaje("Debe escanear un BAC destino", TipoMensaje.MENSAJE_Grave);
            return;
        }

        if (!int.TryParse(txtCantidad.Text, out int cantidad) || cantidad <= 0)
        {
            MostrarMensaje("La cantidad debe ser mayor a 0", TipoMensaje.MENSAJE_Grave);
            return;
        }

        try
        {
            if (_dataAccess == null) return;

            // TODO: Implement article distribution to BAC
            // This requires additional stored procedures not present in the VB.NET version
            // The VB.NET original doesn't have InsertaDetalleBac accessible here
            // Would need: _dataAccess.InsertaDetalleBac(_bacDestino, _grupoBAC, _tablillaBAC, articuloId, cantidad, GlobalData.Usuario.Id);
            
            MostrarMensaje($"Funcionalidad pendiente: Repartir {cantidad} unidades del artículo {_articuloActual} al BAC {_bacDestino}.\nRequiere implementación de stored procedures adicionales.", TipoMensaje.MENSAJE_Informativo);

            // Limpiar para siguiente operación
            _articuloActual = "";
            _bacDestino = "";
            _grupoBAC = 0;
            _tablillaBAC = 0;

            MainThread.BeginInvokeOnMainThread(() =>
            {
                lblArticulo.Text = "";
                lblNombreArticulo.Text = "";
                lblEAN13.Text = "";
                lblPeso.Text = "";
                lblVolumen.Text = "";
                lblBACDestino.Text = "";
                lblGrupoBAC.Text = "";
                lblTablillaBAC.Text = "";
                txtBAC.Text = "";
                txtCantidad.Text = "1";
                txtLecturaCodigo.Focus();
            });
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }
    }

    private void LimpiarDatos()
    {
        _articuloActual = "";
        _bacDestino = "";
        _grupoBAC = 0;
        _tablillaBAC = 0;

        MainThread.BeginInvokeOnMainThread(() =>
        {
            lblArticulo.Text = "";
            lblNombreArticulo.Text = "";
            lblEAN13.Text = "";
            lblPeso.Text = "";
            lblVolumen.Text = "";
            lblBACDestino.Text = "";
            lblGrupoBAC.Text = "";
            lblTablillaBAC.Text = "";
            txtBAC.Text = "";
            txtCantidad.Text = "1";
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
