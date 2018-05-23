using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections;
using System.Data.SqlClient;
using System.Configuration;

namespace ContractSender.Clases
{    
        public class DBAccessWin
        {
            private string _descripcionError = string.Empty;
            private int _noError = 0;
            private bool _error = false;

            public string DescripcionError
            {
                get
                {
                    return _descripcionError;
                }
            }

            public int NoError
            {
                get
                {
                    return _noError;
                }
            }

            public bool Error
            {
                get
                {
                    return _error;
                }
            }
        /// <summary>
        /// Ejecuta un Stored-Procedure existente en la Base de Datos
        /// </summary>
        /// <param name="storedProcedureName"></param>
        /// <param name="paremeters"></param>
        /// <returns></returns>

        public DataTable EjecutarSQLStoredProcedure(string storedProcedureName, ArrayList paremeters)
            {

                SqlConnection _conexion = new SqlConnection(ConfigurationManager.ConnectionStrings["KalaConnection"].ConnectionString);

                SqlCommand com = new SqlCommand();
                try
                {
                    //CrearConexion();
                    _conexion.Open();
                    com.CommandType = CommandType.StoredProcedure;
                    com.CommandText = storedProcedureName;
                    com.Connection = _conexion;
                    com.CommandTimeout = 90000;
                    SqlCommandBuilder.DeriveParameters(com);

                    int indiceParametro = 1;
                    foreach (object param in paremeters)
                    {
                        com.Parameters[indiceParametro].Value = param;
                        indiceParametro++;
                    }

                    SqlDataAdapter adaptador = new SqlDataAdapter(com);
                    DataSet ds = new DataSet();
                    DataTable tabla = new DataTable();

                    adaptador.Fill(tabla);
                    return tabla;

                }
                catch (SqlException ex)
                {
                    this._descripcionError = ex.Message;
                    this._noError = ex.Number;
                    this._error = true;
                    return null;
                }
                finally
                {
                    _conexion.Close();
                }
            }


        }
}

