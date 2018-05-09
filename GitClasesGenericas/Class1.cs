//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Data.SqlClient;
//using System.Collections;
//using System.Data;

//namespace GitClasesGenericas
//{


//    public class BaseDatos
//    {
//        public string CadenaConexion;
//        private ArrayList Transacciones;
//        private SqlConnection NC1 = new SqlConnection();
//        private SqlTransaction Trans;
//        private bool TransOpen;
//        public BaseDatos(string CadenaConexion, bool pvb_ActualizarConexion = true)
//        {
//            NC1.ConnectionString = CadenaConexion;
//            if (pvb_ActualizarConexion)
//                try
//                {
//                    DataTable vlo_Datos = new DataTable("Tmp");
//                    vlo_Datos.Columns.Add("Tiempo");
//                    vlo_Datos.Rows.Add(DateTime.Now.ToString("yyyyMMdd hh:mm:ss tt"));
//                    vlo_Datos.WriteXml("Tiempo.xml");
//                }
//                catch
//                {
//                }
//        }

//        public BaseDatos()
//        {
//        }

//        public bool ConectarDB()
//        {
//            try
//            {
//                NC1.Open();
//            }
//            catch (Exception ex)
//            {
//                return false;
//            }

//            return true;
//        }

//        public bool DesconectarDB()
//        {
//            try
//            {
//                NC1.Close();
//            }
//            catch (Exception ex)
//            {                
//                //MessageBox(ex.Message);
//                return false;
//            }

//            return true;
//        }

//        public bool OpenTrans()
//        {
//            try
//            {
//                if (NC1.State != ConnectionState.Open)
//                    NC1.Open();
//                Trans = NC1.BeginTransaction();
//                TransOpen = true;
//            }
//            catch (Exception ex)
//            {
//                TransOpen = false;
//                //MsgBox(ex.Message);
//                return false;
//            }

//            return true;
//        }

//        public bool CloseTrans()
//        {
//            try
//            {
//                Trans.Commit();
//                NC1.Close();
//            }
//            catch (Exception ex)
//            {
//                TransOpen = false;
//                //MsgBox(ex.Message);
//                return false;
//            }

//            TransOpen = false;
//            return true;
//        }

//        public void RollBackTrans()
//        {
//            try
//            {
//                Trans.Rollback();
//                TransOpen = false;
//            }
//            catch (Exception ex)
//            {
//            }

//            TransOpen = false;
//        }

//        public DLLTipos.GENTipos.Terror EjecutarSQL(SqlCommand CMD, bool pvc_Serializable = false)
//        {
//            if (TransOpen)
//                return EjecutarSQL_T(CMD, pvc_Serializable);
//            else
//                return EjecutarSQL_N(CMD, pvc_Serializable);
//        }

//        private DLLTipos.GENTipos.Terror EjecutarSQL_N(SqlCommand CMD, bool pvc_Serializable)
//        {
//            DLLTipos.GENTipos.Terror control = new DLLTipos.GENTipos.Terror();
//            control.estado = true;
//            control.nerror = 0;
//            control.mensaje = "";
//            try
//            {
//                if (!ConectarDB())
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                    return control;
//                }

//                if (pvc_Serializable)
//                    CMD.CommandType = CommandType.StoredProcedure;
//                CMD.Connection = NC1;
//                CMD.ExecuteNonQuery();
//                if (!DesconectarDB())
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                }
//            }
//            catch (Exception ex)
//            {
//                NC1.Close();
//                control.estado = false;
//                control.nerror = Err.Number;
//                control.mensaje = Err.Description;
//            }

//            return control;
//        }

//        private DLLTipos.GENTipos.Terror EjecutarSQL_T(SqlCommand CMD, bool pvc_Serializable)
//        {
//            DLLTipos.GENTipos.Terror control = new DLLTipos.GENTipos.Terror();
//            control.estado = true;
//            control.nerror = 0;
//            control.mensaje = "";
//            try
//            {
//                if (pvc_Serializable)
//                    CMD.CommandType = CommandType.StoredProcedure;
//                CMD.Connection = NC1;
//                CMD.Transaction = Trans;
//                CMD.ExecuteNonQuery();
//            }
//            catch (Exception ex)
//            {
//                try
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                    RollBackTrans();
//                }
//                catch
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                }
//            }

//            CMD.Dispose();
//            return control;
//        }

//        public DLLTipos.GENTipos.Terror EjecutaSQL(SqlCommand pvo_Comando, bool pvb_StoreProcedure = false)
//        {
//            if (TransOpen)
//                return EjecutaSQL_T(pvo_Comando, pvb_StoreProcedure);
//            else
//                return EjecutaSQL_N(pvo_Comando, pvb_StoreProcedure);
//        }

//        private DLLTipos.GENTipos.Terror EjecutaSQL_T(SqlCommand pvo_Comando, bool pvb_StoreProcedure)
//        {
//            DLLTipos.GENTipos.Terror control = new DLLTipos.GENTipos.Terror();
//            control.estado = true;
//            control.nerror = 0;
//            control.mensaje = "";
//            try
//            {
//                pvo_Comando.Connection = NC1;
//                pvo_Comando.Transaction = Trans;
//                if ((pvb_StoreProcedure))
//                    pvo_Comando.CommandType = CommandType.StoredProcedure;
//                pvo_Comando.ExecuteNonQuery();
//            }
//            catch (Exception vlo_Error)
//            {
//                RollBackTrans();
//                control.estado = false;
//                control.nerror = Err.Number;
//                control.mensaje = Err.Description;
//            }

//            return control;
//        }

//        private DLLTipos.GENTipos.Terror EjecutaSQL_N(SqlCommand pvo_Comando, bool pvb_StoreProcedure)
//        {
//            DLLTipos.GENTipos.Terror control = new DLLTipos.GENTipos.Terror();
//            control.estado = true;
//            control.nerror = 0;
//            control.mensaje = "";
//            try
//            {
//                if (!ConectarDB())
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                    return control;
//                }

//                if ((pvb_StoreProcedure))
//                    pvo_Comando.CommandType = CommandType.StoredProcedure;
//                pvo_Comando.Connection = NC1;
//                pvo_Comando.ExecuteNonQuery();
//                DesconectarDB();
//            }
//            catch (Exception ex)
//            {
//                control.estado = false;
//                control.nerror = Err.Number;
//                control.mensaje = Err.Description;
//            }

//            return control;
//        }

//        public DLLTipos.GENTipos.Terror EjecutaSQL(string SQL, bool getconsecutivo = false)
//        {
//            if (TransOpen)
//                return EjecutaSQL_T(SQL, getconsecutivo);
//            else
//                return EjecutaSQL_N(SQL, getconsecutivo);
//        }

//        private DLLTipos.GENTipos.Terror EjecutaSQL_N(string SQL, bool getconsecutivo)
//        {
//            DLLTipos.GENTipos.Terror control = new DLLTipos.GENTipos.Terror();
//            control.estado = true;
//            control.nerror = 0;
//            control.mensaje = "";
//            try
//            {
//                if (!ConectarDB())
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                }

//                SqlCommand Comando = new SqlCommand(SQL, NC1);
//                Comando.ExecuteNonQuery();
//                if (!DesconectarDB())
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                }
//            }
//            catch (Exception ex)
//            {
//                NC1.Close();
//                control.estado = false;
//                control.nerror = Err.Number;
//                control.mensaje = Err.Description;
//            }

//            return control;
//        }

//        private DLLTipos.GENTipos.Terror EjecutaSQL_T(string sql, bool getconsecutivo)
//        {
//            SqlCommand Comando = new SqlCommand(sql, NC1);
//            DLLTipos.GENTipos.Terror control = new DLLTipos.GENTipos.Terror();
//            control.estado = true;
//            control.nerror = 0;
//            control.mensaje = "";
//            try
//            {
//                Comando.Connection = NC1;
//                Comando.Transaction = Trans;
//                if (getconsecutivo == false)
//                    Comando.ExecuteNonQuery();
//                else
//                    control.tag = Comando.ExecuteScalar.ToString;
//            }
//            catch (Exception ex)
//            {
//                try
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                    RollBackTrans();
//                }
//                catch
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                }
//            }

//            Comando.Dispose();
//            return control;
//        }

//        public bool LlenarDS(ref DataSet DS, SqlCommand CMD, string Nombre = "NoDefinido")
//        {
//            if (TransOpen)
//                return LlenarDS_T(ref DS, CMD, Nombre);
//            else
//                return LlenarDS_N(ref DS, CMD, Nombre);
//        }

//        private bool LlenarDS_N(ref DataSet DS, SqlCommand CMD, string Nombre = "NoDefinido")
//        {
//            try
//            {
//                if (!ConectarDB())
//                    return false;
//                CMD.Connection = NC1;
//                if (Nombre == "DatosStore")
//                    CMD.CommandType = CommandType.StoredProcedure;
//                CMD.CommandTimeout = 0;
//                SqlDataAdapter Adaptador = new SqlDataAdapter(CMD);
//                Adaptador.Fill(DS, Nombre);
//                if (!DesconectarDB())
//                    return false;
//            }
//            catch (Exception ex)
//            {
//                DesconectarDB();
//                //MsgBox(ex.Message);
//                return false;
//            }

//            return true;
//        }

//        private bool LlenarDS_T(ref DataSet DS, SqlCommand CMD, string Nombre = "NoDefinido")
//        {
//            try
//            {
//                CMD.Connection = NC1;
//                CMD.Transaction = Trans;
//                if (Nombre == "DatosStore")
//                    CMD.CommandType = CommandType.StoredProcedure;
//                CMD.CommandTimeout = 0;
//                CMD.ExecuteNonQuery();
//                SqlDataAdapter adaptador = new SqlDataAdapter(CMD);
//                adaptador.Fill(DS, Nombre);
//            }
//            catch (Exception ex)
//            {
//                //MsgBox(ex.Message);
//                return false;
//            }

//            return true;
//        }

//        public bool LlenarDS(ref DataSet DS, string SQL, string Nombre = "NoDefinido")
//        {
//            if (TransOpen)
//                return LlenarDS_T(ref DS, SQL, Nombre);
//            else
//                return LlenarDS_N(ref DS, SQL, Nombre);
//        }

//        private bool LlenarDS_T(ref DataSet DS, string SQL, string Nombre = "NoDefinido")
//        {
//            try
//            {
//                SqlCommand CommandoSQL = new SqlCommand(SQL, NC1, Trans);
//                CommandoSQL.CommandTimeout = 0;
//                CommandoSQL.ExecuteNonQuery();
//                SqlDataAdapter adaptador = new SqlDataAdapter(CommandoSQL);
//                adaptador.Fill(DS, Nombre);
//            }
//            catch (Exception ex)
//            {
//                //MsgBox(ex.Message);
//                return false;
//            }

//            return true;
//        }

//        private bool LlenarDS_N(ref DataSet DS, string SQL, string Nombre = "NoDefinido")
//        {
//            try
//            {
//                if (!ConectarDB())
//                    return false;
//                SqlCommand CommandoSQL = new SqlCommand(SQL, NC1, Trans);
//                CommandoSQL.CommandTimeout = 0;
//                CommandoSQL.ExecuteNonQuery();
//                SqlDataAdapter adaptador = new SqlDataAdapter(CommandoSQL);
//                adaptador.Fill(DS, Nombre);
//                if (!DesconectarDB())
//                    return false;
//            }
//            catch (Exception ex)
//            {
//                DesconectarDB();
//                //MsgBox(ex.Message);
//                return false;
//            }

//            return true;
//        }

//        public bool LlenarDR(ref System.Data.SqlClient.SqlDataReader DR, string SQL, string Nombre = "NoDefinido")
//        {
//            try
//            {
//                SqlClient.SqlCommand Comando = new SqlClient.SqlCommand(SQL);
//                Comando.Connection = NC1;
//                DR = Comando.ExecuteReader;
//            }
//            catch (Exception ex)
//            {
//                //MsgBox(ex.Message);
//                return false;
//            }

//            return true;
//        }

//        public DataSet LLenaDataSet(string pSQL, string pNombre = "NoDefinido")
//        {
//            DataSet DS = new DataSet();
//            LlenarDS(ref DS, pSQL, pNombre);
//            return DS;
//        }

//        public DLLTipos.GENTipos.Terror LlenarDataSet(SqlCommand pvo_Comando, string pvc_Nombre, bool pvb_StoreProcedure = false)
//        {
//            if ((TransOpen))
//                return LlenarDS_Transaccional(pvo_Comando, pvc_Nombre, pvb_StoreProcedure);
//            else
//                return LlenarDS_NoTransaccional(pvo_Comando, pvc_Nombre, pvb_StoreProcedure);
//        }

//        private DLLTipos.GENTipos.Terror LlenarDS_Transaccional(SqlCommand pvo_Comando, string pvc_Nombre, bool pvb_StoreProcedure)
//        {
//            DataSet vlo_Datos = new DataSet();
//            SqlDataAdapter vlo_Adaptador = new SqlDataAdapter();
//            DLLTipos.GENTipos.Terror control = new DLLTipos.GENTipos.Terror();
//            control.estado = true;
//            control.nerror = 0;
//            control.mensaje = "";
//            try
//            {
//                pvo_Comando.Connection = NC1;
//                pvo_Comando.Transaction = Trans;
//                if ((pvb_StoreProcedure))
//                    pvo_Comando.CommandType = CommandType.StoredProcedure;
//                pvo_Comando.ExecuteNonQuery();
//                vlo_Adaptador = new SqlDataAdapter(pvo_Comando);
//                vlo_Adaptador.Fill(vlo_Datos, pvc_Nombre);
//            }
//            catch (Exception ex)
//            {
//                RollBackTrans();
//                control.estado = false;
//                control.mensaje = ex.Message;
//            }

//            control.vlo_DataSet = vlo_Datos;
//            return control;
//        }

//        private DLLTipos.GENTipos.Terror LlenarDS_NoTransaccional(SqlCommand pvo_Comando, string pvc_Nombre, bool pvb_StoreProcedure)
//        {
//            SqlDataAdapter vlo_Adaptador = new SqlDataAdapter();
//            DataSet vlo_Datos = new DataSet();
//            DLLTipos.GENTipos.Terror control = new DLLTipos.GENTipos.Terror();
//            control.estado = true;
//            control.nerror = 0;
//            control.mensaje = "";
//            try
//            {
//                if (!ConectarDB())
//                {
//                    control.estado = false;
//                    control.nerror = Err.Number;
//                    control.mensaje = Err.Description;
//                    return control;
//                }

//                pvo_Comando.Connection = NC1;
//                if ((pvb_StoreProcedure))
//                    pvo_Comando.CommandType = CommandType.StoredProcedure;
//                vlo_Adaptador = new SqlDataAdapter(pvo_Comando);
//                vlo_Adaptador.Fill(vlo_Datos, pvc_Nombre);
//                DesconectarDB();
//            }
//            catch (Exception ex)
//            {
//                DesconectarDB();
//                control.estado = false;
//                control.mensaje = ex.Message;
//            }

//            control.vlo_DataSet = vlo_Datos;
//            return control;
//        }

//        public DataTable LlenaDataTable(string pSQL, string pNombre = "NoDefinido")
//        {
//            try
//            {
//                DataSet DS = new DataSet();
//                if (LlenarDS(ref DS, pSQL, pNombre) == false)
//                {
//                    DataTable vlo_tabla = new DataTable(pNombre);
//                    DS.Tables.Add(vlo_tabla);
//                }

//                return DS.Tables(pNombre).Copy;
//            }
//            catch (Exception ex)
//            {
//                MsgBox(ex.Message);
//                RollBackTrans();
//                return null /* TODO Change to default(_) if this is not a reference type */;
//            }
//        }

//        public bool DatoRepetido(string Tabla, string campo, string Valor, string tipo, string condicion, bool Ind_Trans = false)
//        {
//            string SQL, aux;
//            DataSet DS = new DataSet();
//            aux = IIf(UCase(Trim(tipo)) != "N", "'", "");
//            if (condicion == "")
//                SQL = "select * from " + Tabla + " where " + campo + " = " + aux + Valor + aux + " ";
//            else
//                SQL = "select * from " + Tabla + " where " + condicion;
//            if (Ind_Trans)
//                LlenarDS(DS, SQL);
//            else
//                LlenarDS_T(DS, SQL);
//            if (DS.Tables(0).Rows.Count > 0)
//                return true;
//        }

//        ~BaseDatos()
//        {
//            base.Finalize();
//        }
//    }
//}