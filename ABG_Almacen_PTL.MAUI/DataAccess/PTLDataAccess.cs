//*****************************************************************************************
// PTLDataAccess.cs
// Clase de acceso a datos para operaciones PTL (Pick-to-Light)
// Converted from VB.NET EntornoDeDatos to C# for .NET MAUI
//*****************************************************************************************

using System.Data;
using Microsoft.Data.SqlClient;
using ABG_Almacen_PTL.MAUI.Modules;

namespace ABG_Almacen_PTL.MAUI.DataAccess
{
    public class PTLDataAccess : IDisposable
    {
        private SqlConnection? _connection;
        private bool _disposed = false;

        // Cadena de conexión
        public string ConnectionString { get; set; }

        // Constructor
        public PTLDataAccess()
        {
            // Usar la cadena de conexión global si está disponible
            ConnectionString = string.IsNullOrEmpty(GlobalData.ConexionGestionAlmacen) 
                ? GlobalData.ConexionConfig 
                : GlobalData.ConexionGestionAlmacen;
        }

        // Constructor con cadena de conexión
        public PTLDataAccess(string connectionString)
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

            if (_connection == null)
            {
                _connection = new SqlConnection(ConnectionString);
            }
            else if (_connection.ConnectionString != ConnectionString)
            {
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

        // Asegurar que la conexión esté abierta
        private void EnsureConnectionOpen()
        {
            if (_connection == null || _connection.State != ConnectionState.Open)
            {
                Open();
            }
        }

        // Ejecutar stored procedure y devolver DataTable
        private DataTable ExecuteStoredProcedure(string procedureName, params SqlParameter[] parameters)
        {
            EnsureConnectionOpen();
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

        // ===== Métodos específicos para PTL =====

        // Dame datos BAC de PTL
        public DataTable DameDatosBACdePTL(string bac)
        {
            return ExecuteStoredProcedure("dbo.DameDatosBACdePTL",
                new SqlParameter("@BAC", bac));
        }

        // Consulta BAC de PTL (existencia en GAUBIBAC)
        public DataTable ConsultaBACdePTL(string bac)
        {
            return ExecuteStoredProcedure("dbo.ConsultaBACdePTL",
                new SqlParameter("@BAC", bac));
        }

        // Dame datos de ubicación PTL
        public DataTable DameDatosUbicacionPTL(int alf, int alm, int blo, int fil, int alt)
        {
            return ExecuteStoredProcedure("dbo.DameDatosUbicacionPTL",
                new SqlParameter("@ALF", alf),
                new SqlParameter("@ALM", alm),
                new SqlParameter("@BLO", blo),
                new SqlParameter("@FIL", fil),
                new SqlParameter("@ALT", alt));
        }

        // Ubicar BAC en PTL
        public void UbicarBACenPTL(string bac, int ubicacion, int usuario, out int retorno, out string msgSalida)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.UbicarBACenPTL", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@BAC", bac);
                cmd.Parameters.AddWithValue("@Ubicacion", ubicacion);
                cmd.Parameters.AddWithValue("@Usuario", usuario);

                var paramRetorno = new SqlParameter("@Retorno", SqlDbType.SmallInt);
                paramRetorno.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramRetorno);

                var paramMsg = new SqlParameter("@msgSalida", SqlDbType.VarChar, 1024);
                paramMsg.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramMsg);

                cmd.ExecuteNonQuery();

                retorno = paramRetorno.Value != DBNull.Value ? (short)paramRetorno.Value : -1;
                msgSalida = paramMsg.Value != DBNull.Value ? paramMsg.Value.ToString() ?? "" : "";
            }
        }

        // Cambiar estado BAC de PTL
        public void CambiaEstadoBACdePTL(string bac, int estado, int usuario, out int retorno, out string msgSalida)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.CambiaEstadoBACdePTL", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@BAC", bac);
                cmd.Parameters.AddWithValue("@Estado", estado);
                cmd.Parameters.AddWithValue("@Usuario", usuario);

                var paramRetorno = new SqlParameter("@Retorno", SqlDbType.SmallInt);
                paramRetorno.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramRetorno);

                var paramMsg = new SqlParameter("@msgSalida", SqlDbType.VarChar, 1024);
                paramMsg.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramMsg);

                cmd.ExecuteNonQuery();

                retorno = paramRetorno.Value != DBNull.Value ? (short)paramRetorno.Value : -1;
                msgSalida = paramMsg.Value != DBNull.Value ? paramMsg.Value.ToString() ?? "" : "";
            }
        }

        // Retirar BAC de PTL
        public void RetirarBACdePTL(string bac, int usuario, out int retorno, out string msgSalida)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.RetirarBACdePTL", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@BAC", bac);
                cmd.Parameters.AddWithValue("@Usuario", usuario);

                var paramRetorno = new SqlParameter("@Retorno", SqlDbType.SmallInt);
                paramRetorno.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramRetorno);

                var paramMsg = new SqlParameter("@msgSalida", SqlDbType.VarChar, 1024);
                paramMsg.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramMsg);

                cmd.ExecuteNonQuery();

                retorno = paramRetorno.Value != DBNull.Value ? (short)paramRetorno.Value : -1;
                msgSalida = paramMsg.Value != DBNull.Value ? paramMsg.Value.ToString() ?? "" : "";
            }
        }

        // Dame contenido BAC de Grupo
        public DataTable DameContenidoBacGrupo(int grupo, string bac)
        {
            return ExecuteStoredProcedure("dbo.DameContenidoBacGrupo",
                new SqlParameter("@Grupo", grupo),
                new SqlParameter("@BAC", bac));
        }

        // Dame datos CAJA de PTL
        public DataTable DameDatosCAJAdePTL(string sscc)
        {
            return ExecuteStoredProcedure("dbo.DameDatosCAJAdePTL",
                new SqlParameter("@SSCC", sscc));
        }

        // Dame contenido Caja de Grupo
        public DataTable DameContenidoCajaGrupo(int grupo, int tablilla, string caja)
        {
            return ExecuteStoredProcedure("dbo.DameContenidoCajaGrupo",
                new SqlParameter("@Grupo", grupo),
                new SqlParameter("@Tablilla", tablilla),
                new SqlParameter("@Caja", caja));
        }

        // Dame tipos de cajas activas
        public DataTable DameTiposCajasActivas()
        {
            return ExecuteStoredProcedure("dbo.DameTiposCajasActivas");
        }

        // Dame cajas Grupo Tablilla PTL
        public DataTable DameCajasGrupoTablillaPTL(int grupo, int tablilla)
        {
            return ExecuteStoredProcedure("dbo.DameCajasGrupoTablillaPTL",
                new SqlParameter("@Grupo", grupo),
                new SqlParameter("@Tablilla", tablilla));
        }

        // Dame caja Grupo Tablilla PTL
        public DataTable DameCajaGrupoTablillaPTL(int grupo, int tablilla, string caja)
        {
            return ExecuteStoredProcedure("dbo.DameCajaGrupoTablillaPTL",
                new SqlParameter("@Grupo", grupo),
                new SqlParameter("@Tablilla", tablilla),
                new SqlParameter("@Caja", caja));
        }

        // Crear caja Grupo Tablilla PTL
        public void CrearCajaGrupoTablillaPTL(int grupo, int tablilla, string caja, int tipoCaja, string sscc, string bac)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.CrearCajaGrupoTablillaPTL", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@Grupo", grupo);
                cmd.Parameters.AddWithValue("@Tablilla", tablilla);
                cmd.Parameters.AddWithValue("@Caja", caja);
                cmd.Parameters.AddWithValue("@TipoCaja", tipoCaja);
                cmd.Parameters.AddWithValue("@SSCC", sscc);
                cmd.Parameters.AddWithValue("@BAC", bac);
                cmd.ExecuteNonQuery();
            }
        }

        // Actualiza caja BAC PTL
        public void ActualizaCajaBACPTL(string bac, string caja)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.ActualizaCajaBACPTL", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@BAC", bac);
                cmd.Parameters.AddWithValue("@Caja", caja);
                cmd.ExecuteNonQuery();
            }
        }

        // Traspasa BAC a CAJA de PTL (by ref for SSCC)
        public void TraspasaBACaCAJAdePTLByRef(string bac, int usuario, string sscc, out int retorno, out string msgSalida)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.TraspasaBACaCAJAdePTL", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@BAC", bac);
                cmd.Parameters.AddWithValue("@Usuario", usuario);
                cmd.Parameters.AddWithValue("@SSCC", sscc);

                var paramRetorno = new SqlParameter("@Retorno", SqlDbType.SmallInt);
                paramRetorno.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramRetorno);

                var paramMsg = new SqlParameter("@msgSalida", SqlDbType.VarChar, 1024);
                paramMsg.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramMsg);

                cmd.ExecuteNonQuery();

                retorno = paramRetorno.Value != DBNull.Value ? (short)paramRetorno.Value : -1;
                msgSalida = paramMsg.Value != DBNull.Value ? paramMsg.Value.ToString() ?? "" : "";
            }
        }

        // Dame última caja de BAC
        public DataTable DameUltimaCajaDeBAC(string bac)
        {
            return ExecuteStoredProcedure("dbo.DameUltimaCajaDeBAC",
                new SqlParameter("@BAC", bac));
        }

        // Cambiar tipo de caja PTL
        public void CambiaTipoCajaPTL(int tipoCaja, string bac, string sscc, int usuario)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.CambiaTipoCajaPTL", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@TipoCaja", tipoCaja);
                cmd.Parameters.AddWithValue("@BAC", bac);
                cmd.Parameters.AddWithValue("@SSCC", sscc);
                cmd.Parameters.AddWithValue("@Usuario", usuario);
                cmd.ExecuteNonQuery();
            }
        }

        // Combinar cajas PTL
        public void CombinarCajasPTL(string sscc1, string sscc2, int usuario, out int retorno, out string msgSalida)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.CombinarCajasPTL", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@SSCC1", sscc1);
                cmd.Parameters.AddWithValue("@SSCC2", sscc2);
                cmd.Parameters.AddWithValue("@Usuario", usuario);

                var paramRetorno = new SqlParameter("@Retorno", SqlDbType.SmallInt);
                paramRetorno.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramRetorno);

                var paramMsg = new SqlParameter("@msgSalida", SqlDbType.VarChar, 1024);
                paramMsg.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramMsg);

                cmd.ExecuteNonQuery();

                retorno = paramRetorno.Value != DBNull.Value ? (short)paramRetorno.Value : -1;
                msgSalida = paramMsg.Value != DBNull.Value ? paramMsg.Value.ToString() ?? "" : "";
            }
        }

        // Cambiar unidades artículo en caja PTL
        public void CambiaUnidadesArtCajaPTL(string sscc, int articulo, int cantidad, int usuario, out int retorno, out string msgSalida)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.CambiaUnidadesArtCajaPTL", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@SSCC", sscc);
                cmd.Parameters.AddWithValue("@Articulo", articulo);
                cmd.Parameters.AddWithValue("@Cantidad", cantidad);
                cmd.Parameters.AddWithValue("@Usuario", usuario);

                var paramRetorno = new SqlParameter("@Retorno", SqlDbType.SmallInt);
                paramRetorno.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramRetorno);

                var paramMsg = new SqlParameter("@msgSalida", SqlDbType.VarChar, 1024);
                paramMsg.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramMsg);

                cmd.ExecuteNonQuery();

                retorno = paramRetorno.Value != DBNull.Value ? (short)paramRetorno.Value : -1;
                msgSalida = paramMsg.Value != DBNull.Value ? paramMsg.Value.ToString() ?? "" : "";
            }
        }

        // Dame numerador SSCC Hipódromo
        public DataTable DameNumeradorSSCCHipodromo()
        {
            return ExecuteStoredProcedure("dbo.DameNumeradorSSCCHipodromo");
        }

        // Actualiza numerador SSCC Hipódromo
        public void ActualizaNumeradorSSCCHipodromo(int numerador)
        {
            EnsureConnectionOpen();
            using (var cmd = new SqlCommand("dbo.ActualizaNumeradorSSCCHipodromo", _connection))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Constants.CTE_TiempoEsperaEntornoDatos;
                cmd.Parameters.AddWithValue("@Numerador", numerador);
                cmd.ExecuteNonQuery();
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
    }
}
