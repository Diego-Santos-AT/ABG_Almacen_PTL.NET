using ABG_Almacen_PTL.MAUI.Modules;

namespace ABG_Almacen_PTL.MAUI;

public partial class MainPage : ContentPage
{
	public MainPage()
	{
		InitializeComponent();
		LoadUserInfo();
	}

	protected override void OnAppearing()
	{
		base.OnAppearing();
		LoadUserInfo();
	}

	private void LoadUserInfo()
	{
		lblUsuario.Text = $"Usuario: {GlobalData.Usuario.Nombre}";
		lblEmpresa.Text = $"Empresa: {GlobalData.Empresa}";
		lblPuesto.Text = $"Puesto: {GlobalData.wPuestoTrabajo.Descripcion}";
	}

	private async void OnUbicarBACClicked(object sender, EventArgs e)
	{
		await DisplayAlert("En Desarrollo", "La funcionalidad 'Ubicar BAC' estará disponible próximamente.", "OK");
	}

	private async void OnExtraerBACClicked(object sender, EventArgs e)
	{
		await DisplayAlert("En Desarrollo", "La funcionalidad 'Extraer BAC' estará disponible próximamente.", "OK");
	}

	private async void OnEmpaquetarBACClicked(object sender, EventArgs e)
	{
		await DisplayAlert("En Desarrollo", "La funcionalidad 'Empaquetar BAC' estará disponible próximamente.", "OK");
	}

	private async void OnConsultaPTLClicked(object sender, EventArgs e)
	{
		await DisplayAlert("En Desarrollo", "La funcionalidad 'Consulta PTL' estará disponible próximamente.", "OK");
	}

	private async void OnRepartirArticuloClicked(object sender, EventArgs e)
	{
		await DisplayAlert("En Desarrollo", "La funcionalidad 'Repartir Artículo' estará disponible próximamente.", "OK");
	}

	private async void OnCerrarSesionClicked(object sender, EventArgs e)
	{
		bool result = await DisplayAlert("Cerrar Sesión", 
			"¿Está seguro que desea cerrar la sesión?", 
			"Sí", "No");

		if (result)
		{
			GlobalData.LoginSucceeded = false;
			await Shell.Current.GoToAsync("//LoginPage");
		}
	}
}
