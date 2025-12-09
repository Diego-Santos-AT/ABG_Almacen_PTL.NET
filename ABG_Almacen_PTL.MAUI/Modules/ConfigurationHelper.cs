//*************************************************************************************
// ConfigurationHelper.cs
// Helper class for managing application configuration using MAUI Preferences
// Replaces INI file handling from Windows Forms version
//*************************************************************************************

namespace ABG_Almacen_PTL.MAUI.Modules
{
    public static class ConfigurationHelper
    {
        // Read a configuration value
        public static string ReadConfig(string section, string key, string defaultValue)
        {
            string fullKey = $"{section}_{key}";
            return Preferences.Default.Get(fullKey, defaultValue);
        }

        // Write a configuration value
        public static void WriteConfig(string section, string key, string value)
        {
            string fullKey = $"{section}_{key}";
            Preferences.Default.Set(fullKey, value);
        }

        // Initialize default configuration
        public static void InitializeDefaultConfig()
        {
            // Pantalla
            if (!HasConfig("Pantalla", "MainLeft"))
            {
                WriteConfig("Pantalla", "MainLeft", "-60");
                WriteConfig("Pantalla", "MainTop", "-60");
                WriteConfig("Pantalla", "MainWidth", "15480");
                WriteConfig("Pantalla", "MainHeight", "11220");
            }

            // Conexión
            if (!HasConfig("Conexion", "BDDTime"))
            {
                WriteConfig("Conexion", "BDDTime", "30");
                WriteConfig("Conexion", "BDDConfig", "Config");
                WriteConfig("Conexion", "BDDServ", "SELENE");
                WriteConfig("Conexion", "BDDServLocal", "GROOT");
            }

            // Varios
            if (!HasConfig("Varios", "wDirExport"))
            {
                WriteConfig("Varios", "wDirExport", "");
                WriteConfig("Varios", "UsrDefault", "");
                WriteConfig("Varios", "EmpDefault", "");
                WriteConfig("Varios", "PueDefault", "");
            }
        }

        // Check if a configuration value exists
        public static bool HasConfig(string section, string key)
        {
            string fullKey = $"{section}_{key}";
            return Preferences.Default.ContainsKey(fullKey);
        }

        // Clear all configuration
        public static void ClearAllConfig()
        {
            Preferences.Default.Clear();
        }
    }
}
