//*****************************************************************************************
// ConfigDataAccess.cs
// Clase de acceso a datos para la configuración del sistema
// Converted from VB.NET ADO.NET to C# for .NET MAUI
//*****************************************************************************************

using System.Data;
using Microsoft.Data.SqlClient;
using ABG_Almacen_PTL.MAUI.Modules;

namespace ABG_Almacen_PTL.MAUI.DataAccess
{
    public class ConfigDataAccess : IDisposable
    {
        private SqlConnection? _connection;
        private bool _disposed = false;

        // Cadena de conexión
        public string ConnectionString { get; set; }

        // Constructor
        public ConfigDataAccess()
        {
            ConnectionString = "";
        }

        // Constructor con cadena de conexión
        public ConfigDataAccess(string connectionString)
        {
            ConnectionString = connectionString;
        }

        // Abrir conexión
        public void Open(string? connectionString = null)
        {
            if (!string.IsNullOrEmpty(connectionString))
            {
                ConnectionString = connectionString;
            }
            else if (string.IsNullOrEmpty(ConnectionString))
            {
                // Usar la cadena de conexión de configuración global
                ConnectionString = GlobalData.ConexionConfig;
            }

            if (_connection == null)
            {
                _connection = new SqlConnection(ConnectionString);
            }
            else if (_connection.ConnectionString != ConnectionString)
            {
                // Si la cadena de conexión cambió, cerrar y recrear
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
                _connection.Dispose();
                _connection = new SqlConnection(ConnectionString);
            }

            if (_connection.State != ConnectionState.Open)
            {
                _connection.Open();
            }
        }

        // Cerrar conexión
        public void Close()
        {
            if (_connection != null && _connection.State == ConnectionState.Open)
            {
                _connection.Close();
            }
        }

        // Ejecutar stored procedure
        public DataTable ExecuteStoredProcedure(string procedureName, params SqlParameter[] parameters)
        {
            if (_connection == null || _connection.State != ConnectionState.Open)
            {
                Open();
            }

            var dt = new DataTable();
            using (var cmd = new SqlCommand(procedureName, _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;

                if (parameters != null)
                {
                    cmd.Parameters.AddRange(parameters);
                }

                using (var adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(dt);
                }
            }
            return dt;
        }

        // --- Métodos específicos del sistema ---

        // Busca usuario por nombre
        public DataTable BuscaUsuario(string nombre)
        {
            return ExecuteStoredProcedure("dbo.BuscaUsuario",
                new SqlParameter("@Nombre", nombre));
        }

        // Dame menús del usuario
        public DataTable DameMenusUsuario(string usuario, int empresa)
        {
            return ExecuteStoredProcedure("dbo.DameMenusUsuario",
                new SqlParameter("@Usuario", usuario),
                new SqlParameter("@emp", empresa));
        }

        // Dame usuario por ID
        public DataTable DameUsuarioPorId(int id)
        {
            return ExecuteStoredProcedure("dbo.DameUsuarioPorId",
                new SqlParameter("@Id", id));
        }

        // Dame parámetros de empresa
        public DataTable DameParametrosEmpresa(int empresa)
        {
            return ExecuteStoredProcedure("dbo.DameParametrosEmpresa",
                new SqlParameter("@emp", empresa));
        }

        // Dame código de empresa
        public DataTable DameCodigoEmpresa(string nombre)
        {
            return ExecuteStoredProcedure("dbo.DameCodigoEmpresa",
                new SqlParameter("@nom", nombre));
        }

        // Dame empresas por grupo
        public DataTable DameEmpresasPorGrupo(int grupo)
        {
            return ExecuteStoredProcedure("dbo.DameEmpresasPorGrupo",
                new SqlParameter("@grp", grupo));
        }

        // Dame empresas acceso usuario
        public DataTable DameEmpresasAccesoUsuario(int usuario)
        {
            return ExecuteStoredProcedure("dbo.DameEmpresasAccesoUsuario",
                new SqlParameter("@usr", usuario));
        }

        // Dame grupos por usuario
        public DataTable DameGruposPorUsuario(int userId)
        {
            return ExecuteStoredProcedure("dbo.DameGruposPorUsuario",
                new SqlParameter("@uid", userId));
        }

        // Dame puesto de trabajo
        public DataTable DamePuestoTrabajo(int codigo)
        {
            return ExecuteStoredProcedure("dbo.DamePuestoTrabajo",
                new SqlParameter("@Codigo", codigo));
        }

        // Dame código de puesto
        public DataTable DameCodigoPuesto(string descripcion)
        {
            return ExecuteStoredProcedure("dbo.DameCodigoPuesto",
                new SqlParameter("@DescripcionPuesto", descripcion));
        }

        // Dame puestos
        public DataTable DamePuestos()
        {
            return ExecuteStoredProcedure("dbo.DamePuestos");
        }

        // Dame informes
        public DataTable DameInformes()
        {
            return ExecuteStoredProcedure("dbo.DameInformes");
        }

        // Dame empresa FTP
        public DataTable DameEmpresaFTP(int empresa)
        {
            return ExecuteStoredProcedure("dbo.DameEmpresaFTP",
                new SqlParameter("@Empresa", empresa));
        }

        // Tiene permiso sobre informe
        public int TienePermisoSobreInforme(int usuario, int idInforme)
        {
            if (_connection == null || _connection.State != ConnectionState.Open)
            {
                Open();
            }

            using (var cmd = new SqlCommand("dbo.TienePermisoSobreInforme", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Usuario", usuario);
                cmd.Parameters.AddWithValue("@Id_Informe", idInforme);

                var paramCuantos = new SqlParameter("@Cuantos", SqlDbType.Int)
                {
                    Direction = ParameterDirection.InputOutput,
                    Value = 0
                };
                cmd.Parameters.Add(paramCuantos);

                cmd.ExecuteNonQuery();

                return (int)(paramCuantos.Value ?? 0);
            }
        }

        // Propiedad para acceder a la conexión directamente (compatibilidad)
        public SqlConnection? Config
        {
            get
            {
                if (_connection == null)
                {
                    _connection = new SqlConnection(ConnectionString);
                }
                return _connection;
            }
        }

        // Implementación de IDisposable
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_connection != null)
                    {
                        if (_connection.State == ConnectionState.Open)
                        {
                            _connection.Close();
                        }
                        _connection.Dispose();
                        _connection = null;
                    }
                }
                _disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~ConfigDataAccess()
        {
            Dispose(false);
        }
    }
}
