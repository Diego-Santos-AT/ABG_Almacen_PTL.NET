using ABG_Almacen_PTL.MAUI.DataAccess;
using ABG_Almacen_PTL.MAUI.Modules;
using System.Data;

namespace ABG_Almacen_PTL.MAUI.Pages;

public partial class UbicarBACPage : ContentPage
{
    private PTLDataAccess? _dataAccess;
    private int _ubicacionId = 0;

    public UbicarBACPage()
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

        // Inicializar visualización
        LimpiarDatos();

        try
        {
            switch (codigo.Length)
            {
                case 12: // Unidad de transporte / Ubicación
                    // Comprobar si la lectura es un BAC
                    if (!ValidarBAC(codigo, false))
                    {
                        // Comprobar si la lectura es una ubicación
                        if (!ValidarUbicacion(codigo, false))
                        {
                            // No existe la ubicación / BAC
                            MostrarMensaje("No se ha encontrado Ubicación o BAC", TipoMensaje.MENSAJE_Grave);
                        }
                    }
                    break;
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }

        txtLecturaCodigo.Text = "";
    }

    private bool ValidarBAC(string bac, bool mostrarMensaje)
    {
        if (_dataAccess == null) return false;

        try
        {
            var dtBAC = _dataAccess.DameDatosBACdePTL(bac);

            if (dtBAC.Rows.Count > 0)
            {
                var row = dtBAC.Rows[0];

                double unipes = row["unipes"] != DBNull.Value ? Convert.ToDouble(row["unipes"]) : 0;
                double unipma = row["unipma"] != DBNull.Value ? Convert.ToDouble(row["unipma"]) : 0;
                double univol = row["univol"] != DBNull.Value ? Convert.ToDouble(row["univol"]) : 0;
                double univma = row["univma"] != DBNull.Value ? Convert.ToDouble(row["univma"]) : 0;
                int uninum = row["uninum"] != DBNull.Value ? Convert.ToInt32(row["uninum"]) : 0;
                string unicod = row["unicod"] != DBNull.Value ? row["unicod"].ToString() ?? "" : "";
                int uniest = row["uniest"] != DBNull.Value ? Convert.ToInt32(row["uniest"]) : 0;
                int unigru = row["unigru"] != DBNull.Value ? Convert.ToInt32(row["unigru"]) : 0;
                int unitab = row["unitab"] != DBNull.Value ? Convert.ToInt32(row["unitab"]) : 0;
                string unicaj = row["unicaj"] != DBNull.Value ? row["unicaj"].ToString() ?? "" : "";
                string tipdes = row["tipdes"] != DBNull.Value ? row["tipdes"].ToString() ?? "" : "";

                bool calculoPeso = unipes > unipma;
                bool calculoVolumen = univol > univma;

                // Mostrar los datos
                if (row["ubicod"] == DBNull.Value)
                {
                    RefrescarDatos(0, 0, 0, 0, 0, unicod, uniest, unigru, unitab, unipes, univol, unicaj, tipdes, calculoPeso, calculoVolumen);
                }
                else
                {
                    int ubicod = Convert.ToInt32(row["ubicod"]);
                    int ubialm = row["ubialm"] != DBNull.Value ? Convert.ToInt32(row["ubialm"]) : 0;
                    int ubiblo = row["ubiblo"] != DBNull.Value ? Convert.ToInt32(row["ubiblo"]) : 0;
                    int ubifil = row["ubifil"] != DBNull.Value ? Convert.ToInt32(row["ubifil"]) : 0;
                    int ubialt = row["ubialt"] != DBNull.Value ? Convert.ToInt32(row["ubialt"]) : 0;
                    RefrescarDatos(ubicod, ubialm, ubiblo, ubifil, ubialt, unicod, uniest, unigru, unitab, unipes, univol, unicaj, tipdes, calculoPeso, calculoVolumen);
                }

                if (uninum > 0)
                {
                    MostrarMensaje("El BAC ya se encuentra ubicado", TipoMensaje.MENSAJE_Grave);
                    LimpiarDatos();
                    _ubicacionId = 0;
                }
                else
                {
                    if (_ubicacionId > 0)
                    {
                        // Se ha leído la ubicación anteriormente. Se procede a ubicar el BAC
                        if (UbicarBAC(unicod, _ubicacionId, uniest, rbCerrarBAC.IsChecked))
                        {
                            MostrarMensaje($"Se ha ubicado el BAC: {unicod} en la ubicación de PTL {_ubicacionId}", TipoMensaje.MENSAJE_Exclamacion);
                            _ubicacionId = 0;
                            LimpiarDatos();
                        }
                    }
                    // Else: Se queda pendiente de la lectura de la ubicación o de otro BAC
                }

                return true;
            }
            else
            {
                // No se ha encontrado el BAC. Se comprueba si existe la definición en GAUBIBAC
                var dtConsulta = _dataAccess.ConsultaBACdePTL(bac);

                if (dtConsulta.Rows.Count > 0)
                {
                    string ubibac = dtConsulta.Rows[0]["ubibac"] != DBNull.Value ? dtConsulta.Rows[0]["ubibac"].ToString() ?? "" : "";
                    RefrescarDatos(0, 0, 0, 0, 0, ubibac, 0, 0, 0, 0, 0, "", "", false, false);
                    
                    if (_ubicacionId > 0)
                    {
                        // Se ha leído la ubicación anteriormente. Se procede a ubicar el BAC
                        if (UbicarBAC(ubibac, _ubicacionId, 0, rbCerrarBAC.IsChecked))
                        {
                            MostrarMensaje($"Se ha ubicado el BAC: {ubibac} en la ubicación de PTL {_ubicacionId}", TipoMensaje.MENSAJE_Exclamacion);
                            _ubicacionId = 0;
                            LimpiarDatos();
                        }
                    }
                    // Else: Se queda pendiente de la lectura de la ubicación o de otro BAC

                    return true;
                }
                else
                {
                    if (mostrarMensaje)
                        MostrarMensaje("No existe el BAC", TipoMensaje.MENSAJE_Grave);
                }
            }
        }
        catch (Exception ex)
        {
            if (mostrarMensaje)
                MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }

        return false;
    }

    private bool ValidarUbicacion(string ubicacion, bool mostrarMensaje)
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
                _ubicacionId = ubicod;

                // Si existe comprueba si tiene un BAC asociado
                if (row["unicod"] == DBNull.Value)
                {
                    MainThread.BeginInvokeOnMainThread(() =>
                    {
                        lblUbicacion.Text = $"({ubicod}) {iALM:000}.{iBLO:000}.{iFIL:000}.{iALT:000}";
                    });

                    // Si se ha leído el BAC anteriormente. Se procede a ubicar el BAC
                    if (!string.IsNullOrEmpty(lblBAC.Text))
                    {
                        int estadoBAC = lblEstadoBAC.Text == "ABIERTO" ? 0 : 1;
                        if (UbicarBAC(lblBAC.Text, _ubicacionId, estadoBAC, rbCerrarBAC.IsChecked))
                        {
                            MostrarMensaje($"Se ha ubicado el BAC: {lblBAC.Text} en la ubicación de PTL {_ubicacionId}", TipoMensaje.MENSAJE_Exclamacion);
                            _ubicacionId = 0;
                            LimpiarDatos();
                        }
                    }
                    // Else: Se queda pendiente de la lectura del BAC o de otra ubicación
                }
                else
                {
                    MostrarMensaje("La Ubicación ya tiene asociado un BAC", TipoMensaje.MENSAJE_Grave);
                    _ubicacionId = 0;
                }

                return true;
            }
            else
            {
                if (mostrarMensaje)
                    MostrarMensaje("No existe la Unidad de Transporte", TipoMensaje.MENSAJE_Grave);
                
                MainThread.BeginInvokeOnMainThread(() =>
                {
                    lblUbicacion.Text = "";
                });
                _ubicacionId = 0;
            }
        }
        catch (Exception ex)
        {
            if (mostrarMensaje)
                MostrarMensaje($"Error: {ex.Message}", TipoMensaje.MENSAJE_Grave);
        }

        return false;
    }

    private void RefrescarDatos(int codUbicacion, int alm, int blo, int fil, int alt,
                                 string bac, int estadoBAC, int grupo, int tablilla,
                                 double peso, double volumen, string tipoCaja, string nombreCaja,
                                 bool pesoExcedido, bool volumenExcedido)
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            if (codUbicacion == 0)
            {
                lblUbicacion.Text = "SIN UBICACION";
            }
            else
            {
                lblUbicacion.Text = $"({codUbicacion}) {alm:000}.{blo:000}.{fil:000}.{alt:000}";
            }

            lblBAC.Text = bac;
            lblEstadoBAC.Text = estadoBAC == 0 ? "ABIERTO" : "CERRADO";
            lblEstadoBAC.BackgroundColor = estadoBAC == 0 ? Colors.White : Colors.LightGreen;

            lblGrupo.Text = grupo.ToString();
            lblTablilla.Text = tablilla.ToString();
            lblUds.Text = "0";

            lblPeso.Text = peso.ToString("F3");
            lblPeso.BackgroundColor = pesoExcedido ? Colors.LightCoral : Colors.White;

            lblVolumen.Text = volumen.ToString("F3");
            lblVolumen.BackgroundColor = volumenExcedido ? Colors.LightCoral : Colors.White;

            lblTipoCaja.Text = tipoCaja;
            lblNombreCaja.Text = nombreCaja;
        });
    }

    private void LimpiarDatos()
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            lblUbicacion.Text = "";
            lblBAC.Text = "";
            lblEstadoBAC.Text = "";
            lblEstadoBAC.BackgroundColor = Colors.White;
            lblGrupo.Text = "";
            lblTablilla.Text = "";
            lblUds.Text = "";
            lblPeso.Text = "";
            lblPeso.BackgroundColor = Colors.White;
            lblVolumen.Text = "";
            lblVolumen.BackgroundColor = Colors.White;
            lblTipoCaja.Text = "";
            lblNombreCaja.Text = "";
        });
    }

    private bool UbicarBAC(string bac, int ubicacion, int estado, bool estadoFinal)
    {
        if (_dataAccess == null) return false;

        try
        {
            _dataAccess.UbicarBACenPTL(bac, ubicacion, GlobalData.Usuario.Id, out int retorno, out string msgSalida);

            if (retorno == 0)
            {
                // Cambiar estado de BAC si es necesario
                if ((estado == 0) == estadoFinal)
                {
                    int nEstado = estadoFinal ? 1 : 0;
                    if (CambiarEstadoBAC(bac, nEstado))
                    {
                        MainThread.BeginInvokeOnMainThread(() =>
                        {
                            lblEstadoBAC.Text = nEstado == 0 ? "ABIERTO" : "CERRADO";
                            lblEstadoBAC.BackgroundColor = nEstado == 0 ? Colors.White : Colors.LightGreen;
                        });
                    }
                }
                return true;
            }
            else
            {
                MostrarMensaje($"No se ha podido ubicar el BAC en la estantería de PTL. {msgSalida}", TipoMensaje.MENSAJE_Grave);
                return false;
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error al ubicar BAC: {ex.Message}", TipoMensaje.MENSAJE_Grave);
            return false;
        }
    }

    private bool CambiarEstadoBAC(string bac, int estado)
    {
        if (_dataAccess == null) return false;

        try
        {
            _dataAccess.CambiaEstadoBACdePTL(bac, estado, GlobalData.Usuario.Id, out int retorno, out string msgSalida);

            if (retorno == 0)
            {
                return true;
            }
            else
            {
                MostrarMensaje($"No se ha podido cambiar el estado al BAC {msgSalida}", TipoMensaje.MENSAJE_Grave);
                return false;
            }
        }
        catch (Exception ex)
        {
            MostrarMensaje($"Error al cambiar estado: {ex.Message}", TipoMensaje.MENSAJE_Grave);
            return false;
        }
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

    private void OnCancelarClicked(object? sender, EventArgs e)
    {
        LimpiarDatos();
        _ubicacionId = 0;
        txtLecturaCodigo.Focus();
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
