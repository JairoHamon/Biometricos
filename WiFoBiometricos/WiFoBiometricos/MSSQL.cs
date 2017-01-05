using System;
using System.Collections.Generic;
using System.Text;


using System.IO;


using System.Data;
using System.Data.SqlClient;
using System.Configuration;

using System.Configuration;



namespace netveloper.DataManager
{

    public class MSSQL
    {

        protected static SqlConnection MyConn = null;

        #region "Private methods - Data Access Manager"
        //
        private static void AttachParameters(SqlCommand command, SqlParameter[] commandParameters)
        {
            if (command == null) throw new ArgumentNullException("command");
            if (commandParameters != null)
            {
                foreach (SqlParameter p in commandParameters)
                {
                    if (p != null)
                    {
                        // 
                        if ((p.Direction == ParameterDirection.InputOutput ||
                            p.Direction == ParameterDirection.Input) && (p.Value == null))
                        {
                            p.Value = DBNull.Value;
                        }
                        command.Parameters.Add(p);
                    }
                }
            }
        }
        //
        private static void PrepareCommand(SqlCommand command, SqlConnection connection, CommandType commandType, string commandText,
            SqlParameter[] commandParameters)
        {
            if (command == null) throw new ArgumentNullException("command");
            if (commandText == null || commandText.Length == 0) throw new ArgumentNullException("commandText");

            command.Connection = connection;
            command.CommandText = commandText;
            command.CommandType = commandType;

            if (commandParameters != null)
            {
                AttachParameters(command, commandParameters);
            }
            return;
        }
        //
        #endregion

        #region " Métodos públicos Connection y Close "
        //
        public static void Connection(string ConnectionString)
        {
            string sConnection = string.Empty;

            if (ConnectionString == "")
            {
                sConnection = ConfigurationManager.ConnectionStrings["SqlConnectionString"].ToString();
            }
            else
            {
                sConnection = ConfigurationManager.ConnectionStrings[ConnectionString].ToString();
                //sConnection = ConnectionString;
            }

            try
            {
                MyConn = new SqlConnection(sConnection);
                MyConn.Close();
                MyConn.Open();
            }
            catch (Exception ex)
            {
                cargalogsql("Error Conexión: " + ex.Message);
            }
        }
        //
        public static void Close()
        {

            try
            {
                MyConn.Close();
                MyConn.Dispose();
            }
            catch
            {
                throw;
            }
        }
        #endregion

        #region " ExecuteDataSet - ExecuteScalar - ExecuteNonQuery "

        public static DataSet ExecuteDataset(CommandType commandType, string commandText)
        {
            return ExecuteDataset(commandType, commandText, (SqlParameter[])null);
        }
        //
        public static DataSet ExecuteDataset(CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {
            DataSet ds = new DataSet();

            try
            {
                SqlCommand cmd = new SqlCommand();

                PrepareCommand(cmd, MyConn, commandType, commandText, commandParameters);

                SqlDataAdapter da = new SqlDataAdapter(cmd);

                da.Fill(ds);
                cmd.Parameters.Clear();
            }
            catch
            {
                throw;
            }

            return ds;
        }
        //
        public static object ExecuteScalar(CommandType commandType, string commandText)
        {
            return ExecuteScalar(commandType, commandText, (SqlParameter[])null);
        }
        //
        public static object ExecuteScalar(CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {

            object retval = null;

            try
            {
                SqlCommand cmd = new SqlCommand();
                PrepareCommand(cmd, MyConn, commandType, commandText, commandParameters);


                retval = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
            }
            catch
            {
                throw;
            }

            return retval;
        }

        public static int ExecuteNonQuery(CommandType commandType, string commandText)
        {
            return ExecuteNonQuery(commandType, commandText, (SqlParameter[])null);
        }

        public static int ExecuteNonQuery(CommandType commandType, string commandText, params SqlParameter[] commandParameters)
        {

            int retval = 0;
            try
            {

                SqlCommand cmd = new SqlCommand();

                PrepareCommand(cmd, MyConn, commandType, commandText, commandParameters);

                retval = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
            }
            catch
            {
                throw;
            }

            return retval;
        }

        #endregion

        public static void cargalogsql(string mensaje)
        {
            string DirectorioLog = "";
            DirectorioLog = ConfigurationManager.AppSettings["DirectorioLog"];

            try
            {

                if (!System.IO.Directory.Exists(DirectorioLog))
                {
                    System.IO.Directory.CreateDirectory(DirectorioLog);
                }
                string fileName = DirectorioLog + @"\Log" + DateTime.Now.Day.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Year.ToString() + ".txt";
                // esto inserta texto en un archivo existente, si el archivo no existe lo crea
                StreamWriter writer = File.AppendText(fileName);
                writer.WriteLine(mensaje);
                writer.Close();
                writer.Dispose();

            }
            catch
            {
                Console.WriteLine("Error");
            }

        }

    }


}
