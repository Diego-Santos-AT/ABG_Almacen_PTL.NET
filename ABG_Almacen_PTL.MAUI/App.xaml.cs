using Microsoft.Extensions.DependencyInjection;

namespace ABG_Almacen_PTL.MAUI;

public partial class App : Application
{
	public App()
	{
		InitializeComponent();
	}

	protected override Window CreateWindow(IActivationState? activationState)
	{
		var shell = new AppShell();
		// Start with login page
		shell.GoToAsync("//LoginPage");
		return new Window(shell);
	}
}