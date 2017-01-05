using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net.Mail;

using System.Data.OleDb;

using System.Data.SqlClient;
using netveloper.DataManager;

namespace WiFoBiometricos
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        private void label2_Click(object sender, EventArgs e)
        {
            
        }

        public void enviarcorreo(string emaildestino, string CopiaOculta, string Mensaje, string asunto, string adjunto1, string adjunto2)
        {

            String FROM = emaildestino;
            FROM = "seguimiento.biometrico@serviciodeempleo.gov.co";
            String TO = emaildestino;
            String COPIA = CopiaOculta;
            string[] TOs;
            string correo = "";
            int anterior = 0;
            int final = 0;
            try
            {
                String SUBJECT = asunto;
                String BODY = @Mensaje;

                String SMTP_USERNAME = "AKIAJWMRXA3NHDBVMLSQ";  // Replace with your SMTP username. 
                String SMTP_PASSWORD = "AkrhuKpi9FIcbLe4z8oV7lY/Ahtonn++P62nk8scGoF/";  // Replace with your SMTP password.

                String HOST = "email-smtp.us-east-1.amazonaws.com";

                int PORT = 587;//already tried with all recommended ports

                SmtpClient client = new SmtpClient(HOST, PORT);
                System.Net.NetworkCredential mailAuthentication = new System.Net.NetworkCredential(SMTP_USERNAME, SMTP_PASSWORD);
                client.EnableSsl = true;
                client.UseDefaultCredentials = false;
                client.Credentials = mailAuthentication;




                MailMessage mess = new MailMessage();

                for (int i = 0; i < CopiaOculta.Length; i++)
                {
                    if (CopiaOculta.Substring(i, 1) == ";")
                    {
                        correo = CopiaOculta.Substring(anterior, i - anterior);
                        MailAddress bcc = new MailAddress(correo);
                        mess.Bcc.Add(bcc);
                        anterior = i + 1;
                    }
                    final = i;
                }
                /*correo = CopiaOculta.Substring(anterior, final - anterior);
                MailAddress bcc2 = new MailAddress(correo);
                mess.Bcc.Add(bcc2);*/

                        mess.IsBodyHtml = true;
                mess.Body = BODY;
                mess.From = new MailAddress(FROM);
                if (adjunto1 != "")
                {
                    mess.Attachments.Add(new Attachment(adjunto1));
                }
                if (adjunto2 != "")
                {
                    mess.Attachments.Add(new Attachment(adjunto2));
                }
                TOs = TO.Split(';');
                foreach (string i in TOs)
                {
                    mess.To.Add(new MailAddress(i));
                }

                /*if (COPIA != "")
                {
                    mess.CC.Add(new MailAddress(COPIA));
                }*/

                mess.Subject = SUBJECT;



                /* try
                 {*/
                client.Send(mess);
            }
            catch (Exception ex)
            {
                label2.Text = "No se envio el correo, el mensjae que arroja es: " + ex.Message;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string sqlConsulta = "";
            string emaildestino = "";
            string CopiaOculta = "";
            string cuerpo = "";
            string cuerpop = "";
            string cuerpojefe = "";
            string Encabezado = "";

            int cuenta = 1;
            string asunto = "";
            string adjunto1 = "";
            string adjunto2 = "";
            string nombrefuncionario = "";
            string correoa = "";
            string identificacion = "";

            string estiloE = "";
            string estiloS = "";
            string ValorrReader4 = "";
            int agnoini = 2016;
            int mesini = 1;
            int diaini = 1;

            int agnofin = 2016;
            int mesfin = 1;
            int diafin = 1;
            int area = 0;

            int intHace = 7;
            string entro = "NO";
            DateTime fecha = DateTime.Now;
            DateTime PrimeraSalida = DateTime.Now;
            DateTime Entrada = DateTime.Now;
            DateTime Salida = DateTime.Now;

            string campo1, campo2, campo3, campo4;
            campo1 = campo2 = campo3 = campo4 = "";

            agnoini = DateTime.Now.AddDays(-11).Year;
            mesini = DateTime.Now.AddDays(-11).Month;
            diaini = DateTime.Now.AddDays(-11).Day;

            agnofin = DateTime.Now.Year;
            mesfin = DateTime.Now.Month;
            diafin = DateTime.Now.Day;


            /* INFORMACIÓN JEFE */

            TimeSpan DIFERENCIA;
            TimeSpan DiferenciaEntrada;
            TimeSpan DiferenciaSalida;
            TimeSpan DiferenciaPrimeraSalida;

            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();

            try
            {
                if (File.Exists(Path.Combine(@"C:\Biometricos\", "access.mdb")))
                {
                    File.Delete(Path.Combine(@"C:\Biometricos\", "access.mdb"));
                }
                File.Copy(@"\\lws008\ZKAccess3.5\access.mdb", Path.Combine("C:\\Biometricos\\", "access.mdb"));
     

                conn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;" + @"Data source= C:\Biometricos\access.mdb";
                

                DataSet DSAreas = new DataSet();
                MSSQL.Connection("");
                DSAreas = MSSQL.ExecuteDataset(CommandType.StoredProcedure, "ConsultaAreas");
                MSSQL.Close();

                foreach (DataRow DAAreas in DSAreas.Tables[0].Rows)
                {
                    cuerpo = "";
                    cuerpojefe = "";
                    cuerpop = "";
                    identificacion = "";
                    Encabezado = "";
                    area = System.Convert.ToInt16(DAAreas["id"]);
                    identificacion = "'" + DAAreas["Identificacion"].ToString() + "'";
                    correoa = DAAreas["Correo"].ToString();
                    nombrefuncionario = DAAreas["Nombre"].ToString();

                    Encabezado = "<body><p class=\"MsoNormal\">Cordial saludo,<o:p></o:p></p><p class=\"MsoNormal\"><o:p>&nbsp;</o:p>";
                    /* error */
                    //Encabezado = Encabezado + "</p><p class=\"MsoNormal\">Se da alcance al correo anterior, debido a que no se descargó la información de los dispositivos lectores de huella de los días jueves 17 de noviembre parcial y viernes 18 todo el día.<br><br><o:p></o:p></p>";
                    /* error */
                    Encabezado = Encabezado + "</p><p class=\"MsoNormal\">En la siguiente tabla se encuentra la información de sus accesos a la Unidad:<br><br><o:p></o:p></p>";
                    Encabezado = Encabezado + "<table  border=\"1\">";
                    Encabezado = Encabezado + "<tr style=\"background-color: #FF0000; font-family: 'Arial Narrow'; color: #FFFFFF; font-weight: bold\">";
                    Encabezado = Encabezado + "<td align =\"center\">FUNCIONARIO</td>";
                    Encabezado = Encabezado + "<td align =\"center\">FECHA</td>";
                    Encabezado = Encabezado + "<td align =\"center\">HORA INGRESO</td>";
                    Encabezado = Encabezado + "<td align =\"center\">HORA SALIDA</td>";
                    Encabezado = Encabezado + "<td align =\"center\">DURACIÓN</td>";
                    Encabezado = Encabezado + "</tr>";

                    for (int nDia = 0; nDia < 7; nDia++)
                    {
                        agnoini = DateTime.Now.AddDays(nDia - intHace).Year;
                        mesini = DateTime.Now.AddDays(nDia - intHace).Month;
                        diaini = DateTime.Now.AddDays(nDia - intHace).Day;

                        agnofin = DateTime.Now.AddDays(nDia - (intHace - 1)).Year;
                        mesfin = DateTime.Now.AddDays(nDia - (intHace - 1)).Month;
                        diafin = DateTime.Now.AddDays(nDia - (intHace - 1)).Day;

                        sqlConsulta = "select C1.*, c2.horaMaxima";  //, c3.primerasalida from ";
                        sqlConsulta = sqlConsulta + " FROM (SELECT Year(Checkinout.CheckTime) as Agno, Month(Checkinout.CheckTime) as Mes, Day(Checkinout.CheckTime) as dia, userinfo.name, min(Checkinout.CheckTime) AS horaMinima";
                        sqlConsulta = sqlConsulta + " FROM Checkinout ";
                        sqlConsulta = sqlConsulta + " inner join Userinfo ";
                        sqlConsulta = sqlConsulta + " on Checkinout.userid = Userinfo.userid ";
                        sqlConsulta = sqlConsulta + " WHERE  Checkinout.CheckTime Between Format(#" + diaini.ToString() + "/" + mesini.ToString() + "/" + agnoini.ToString() + "#,\"mm / dd / yyyy\") And Format(#" + diafin.ToString() + "/" + mesfin.ToString() + "/" + agnofin.ToString() + "#,\"mm / dd / yyyy\")";
                        sqlConsulta = sqlConsulta + " and USERINFO.[pager] in (" + identificacion + ") ";
                        sqlConsulta = sqlConsulta + " GROUP BY Year(Checkinout.CheckTime), Month(Checkinout.CheckTime), Day(Checkinout.CheckTime), userinfo.name) C1 ";
                        sqlConsulta = sqlConsulta + " inner join ";
                        sqlConsulta = sqlConsulta + " (SELECT Year(Checkinout.CheckTime) as Agno, Month(Checkinout.CheckTime) as Mes, Day(Checkinout.CheckTime) as dia, userinfo.name, max(Checkinout.CheckTime) AS horaMaxima ";
                        sqlConsulta = sqlConsulta + " FROM Checkinout ";
                        sqlConsulta = sqlConsulta + " right join Userinfo ";
                        sqlConsulta = sqlConsulta + " on Checkinout.userid = Userinfo.userid ";
                        sqlConsulta = sqlConsulta + " WHERE  Checkinout.CheckTime Between Format(#" + diaini.ToString() + "/" + mesini.ToString() + "/" + agnoini.ToString() + "#,\"mm / dd / yyyy\") And Format(#" + diafin.ToString() + "/" + mesfin.ToString() + "/" + agnofin.ToString() + "#,\"mm / dd / yyyy\")";
                        sqlConsulta = sqlConsulta + " and USERINFO.[pager] in (" + identificacion + ") ";
                        sqlConsulta = sqlConsulta + " GROUP BY Year(Checkinout.CheckTime), Month(Checkinout.CheckTime), Day(Checkinout.CheckTime), userinfo.name) C2 ";
                        sqlConsulta = sqlConsulta + " on C1.name = C2.name ";
                        sqlConsulta = sqlConsulta + " and C1.Agno = C2.Agno ";
                        sqlConsulta = sqlConsulta + " and C1.Mes = C2.Mes ";
                        sqlConsulta = sqlConsulta + " and C1.dia = C2.dia ";
                        sqlConsulta = sqlConsulta + " ORDER BY C1.name, C1.Agno, C1.Mes, C1.dia ";

                        conn.Open();

                        OleDbCommand comando = new OleDbCommand(sqlConsulta, conn);
                        OleDbDataReader lectura = comando.ExecuteReader();

                        entro = "NO";

                        while (lectura.Read())
                        {
                            cuenta = cuenta + 1;
                            entro = "SI";
                            campo1 = lectura[0].ToString();
                            campo2 = lectura[1].ToString();
                            campo3 = lectura[2].ToString();
                            campo4 = lectura[3].ToString();

                            Entrada = System.Convert.ToDateTime(lectura[4].ToString());
                            Salida = System.Convert.ToDateTime(lectura[5].ToString());

                            estiloE = estiloS = "";
                            DIFERENCIA = System.Convert.ToDateTime(lectura[5].ToString()) - System.Convert.ToDateTime(lectura[4].ToString());

                            if (Entrada.DayOfWeek == DayOfWeek.Friday)
                            {
                                fecha = System.Convert.ToDateTime(lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + " 08:01:00");

                            }
                            else
                            {
                                fecha = System.Convert.ToDateTime(lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + " 07:31:00");
                            }

                            DiferenciaEntrada = Entrada - System.Convert.ToDateTime(fecha);
                            DiferenciaPrimeraSalida = Entrada - PrimeraSalida;

                            if (DiferenciaEntrada.TotalMinutes <= 0)
                            {
                                estiloE = "style=\"color: #008000\"";
                            }
                            else if (DiferenciaEntrada.TotalMinutes <= 15)
                            {
                                estiloE = "style = \"color: #FF9900\"";
                            }
                            else
                            {
                                estiloE = "style=\"color: #FF0000\"";
                            }
                            ValorrReader4 = lectura[4].ToString();

                            fecha = System.Convert.ToDateTime(lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + " 17:00:00");

                            DiferenciaSalida = Salida - System.Convert.ToDateTime(fecha);

                            if (DiferenciaSalida.TotalMinutes < 0)
                            {
                                estiloS = "style=\"color: #FF0000\"";

                            }
                            else if (DiferenciaSalida.TotalMinutes <= 60)
                            {
                                estiloS = "style=\"color: #008000\"";

                            }
                            else
                            {
                                estiloS = "style = \"color: #FF9900\"";
                            }

                            if (DIFERENCIA.TotalMinutes == 0)
                            {
                                estiloS = "style=\"color: #090909\"";
                                cuerpojefe = cuerpojefe + "<TR><TD>" + lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + "</TD>";
                                cuerpojefe = cuerpojefe + "<TD " + estiloE + ">" + ValorrReader4.Substring(10, ValorrReader4.Length - 10) + "</TD>";
                                cuerpojefe = cuerpojefe + "<TD " + estiloS + ">N.R.</TD>";
                                cuerpojefe = cuerpojefe + "<TD>" + DIFERENCIA.ToString() + "</TD></TR>";
                            }
                            else
                            {
                                cuerpojefe = cuerpojefe + "<TR><TD>" + lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + "</TD>";
                                cuerpojefe = cuerpojefe + "<TD " + estiloE + ">" + ValorrReader4.Substring(10, ValorrReader4.Length - 10) + "</TD>";
                                cuerpojefe = cuerpojefe + "<TD " + estiloS + ">" + Salida.ToString().Substring(10, Salida.ToString().Length - 10) + "</TD>";
                                cuerpojefe = cuerpojefe + "<TD>" + DIFERENCIA.ToString() + "</TD></TR>";
                            }
                        }
                        if (entro == "NO")
                        {
                            if (DateTime.Now.AddDays(nDia - intHace).DayOfWeek != DayOfWeek.Saturday)
                            {
                                if (DateTime.Now.AddDays(nDia - intHace).DayOfWeek != DayOfWeek.Sunday)
                                {
                                    cuenta = cuenta + 1;
                                    estiloE = estiloS = "style=\"color: #090909\"";
                                    cuerpojefe = cuerpojefe + "<TR><TD>" + diaini.ToString().PadLeft(2, '0') + "/" + mesini.ToString().PadLeft(2, '0') + "/" + agnoini.ToString() + "</TD>";
                                    cuerpojefe = cuerpojefe + "<TD " + estiloE + ">N.R.</TD>";
                                    cuerpojefe = cuerpojefe + "<TD " + estiloS + ">N.R.</TD>";
                                    cuerpojefe = cuerpojefe + "<TD>N.R.</TD></TR>";
                                }
                            }
                        }
                        conn.Close();

                    }
                    cuerpojefe = "<TR><TD rowspan = " + cuenta.ToString() + ">" + nombrefuncionario + "</TD>" + cuerpojefe;
                    cuerpojefe = cuerpojefe + "</TR></table><BR><BR><BR>";
                    cuerpojefe = cuerpojefe + "</p><p class=\"MsoNormal\">A continuación se encuenta los accesos a la Unidad de los funcionarios que estan a su cargo:<br><br><o:p></o:p></p>";

                    cuerpojefe = cuerpojefe + "<table  border=\"1\">";
                    cuerpojefe = cuerpojefe + "<tr style=\"background-color: #FF0000; font-family: 'Arial Narrow'; color: #FFFFFF; font-weight: bold\">";
                    cuerpojefe = cuerpojefe + "<td align =\"center\">FUNCIONARIO</td>";
                    cuerpojefe = cuerpojefe + "<td align =\"center\">FECHA</td>";
                    cuerpojefe = cuerpojefe + "<td align =\"center\">HORA INGRESO</td>";
                    cuerpojefe = cuerpojefe + "<td align =\"center\">HORA SALIDA</td>";
                    cuerpojefe = cuerpojefe + "<td align =\"center\">DURACIÓN</td>";
                    cuerpojefe = cuerpojefe + "</tr>";


                    cuerpo = Encabezado + cuerpojefe;



                    /* funcionarios */
                    SqlParameter[] idArea = new SqlParameter[1];
                    idArea[0] = new SqlParameter("@Id", SqlDbType.VarChar, 200);
                    idArea[0].Value = area;

                    DataSet DSCedulas = new DataSet();
                    MSSQL.Connection("");
                    DSCedulas = MSSQL.ExecuteDataset(CommandType.StoredProcedure, "ConsultaCedulas", idArea);
                    MSSQL.Close();
                    cuenta = 1;
                    foreach (DataRow DRCedulas in DSCedulas.Tables[0].Rows)
                    {
                        identificacion = "'" + DRCedulas["Identificacion"].ToString() + "'";
                        nombrefuncionario = DRCedulas["nombre"].ToString();
                        for (int nDia = 0; nDia < 7; nDia++)
                        {

                            agnoini = DateTime.Now.AddDays(nDia - intHace).Year;
                            mesini = DateTime.Now.AddDays(nDia - intHace).Month;
                            diaini = DateTime.Now.AddDays(nDia - intHace).Day;

                            agnofin = DateTime.Now.AddDays(nDia - (intHace - 1)).Year;
                            mesfin = DateTime.Now.AddDays(nDia - (intHace - 1)).Month;
                            diafin = DateTime.Now.AddDays(nDia - (intHace - 1)).Day;

                            sqlConsulta = "select C1.*, c2.horaMaxima";  //, c3.primerasalida from ";
                            sqlConsulta = sqlConsulta + " FROM (SELECT Year(Checkinout.CheckTime) as Agno, Month(Checkinout.CheckTime) as Mes, Day(Checkinout.CheckTime) as dia, userinfo.name, min(Checkinout.CheckTime) AS horaMinima";
                            sqlConsulta = sqlConsulta + " FROM Checkinout ";
                            sqlConsulta = sqlConsulta + " inner join Userinfo ";
                            sqlConsulta = sqlConsulta + " on Checkinout.userid = Userinfo.userid ";
                            sqlConsulta = sqlConsulta + " WHERE  Checkinout.CheckTime Between Format(#" + diaini.ToString() + "/" + mesini.ToString() + "/" + agnoini.ToString() + "#,\"mm / dd / yyyy\") And Format(#" + diafin.ToString() + "/" + mesfin.ToString() + "/" + agnofin.ToString() + "#,\"mm / dd / yyyy\")";
                            sqlConsulta = sqlConsulta + " and USERINFO.[pager] in (" + identificacion + ") ";
                            sqlConsulta = sqlConsulta + " GROUP BY Year(Checkinout.CheckTime), Month(Checkinout.CheckTime), Day(Checkinout.CheckTime), userinfo.name) C1 ";
                            sqlConsulta = sqlConsulta + " inner join ";
                            sqlConsulta = sqlConsulta + " (SELECT Year(Checkinout.CheckTime) as Agno, Month(Checkinout.CheckTime) as Mes, Day(Checkinout.CheckTime) as dia, userinfo.name, max(Checkinout.CheckTime) AS horaMaxima ";
                            sqlConsulta = sqlConsulta + " FROM Checkinout ";
                            sqlConsulta = sqlConsulta + " right join Userinfo ";
                            sqlConsulta = sqlConsulta + " on Checkinout.userid = Userinfo.userid ";
                            sqlConsulta = sqlConsulta + " WHERE  Checkinout.CheckTime Between Format(#" + diaini.ToString() + "/" + mesini.ToString() + "/" + agnoini.ToString() + "#,\"mm / dd / yyyy\") And Format(#" + diafin.ToString() + "/" + mesfin.ToString() + "/" + agnofin.ToString() + "#,\"mm / dd / yyyy\")";
                            sqlConsulta = sqlConsulta + " and USERINFO.[pager] in (" + identificacion + ") ";
                            sqlConsulta = sqlConsulta + " GROUP BY Year(Checkinout.CheckTime), Month(Checkinout.CheckTime), Day(Checkinout.CheckTime), userinfo.name) C2 ";
                            sqlConsulta = sqlConsulta + " on C1.name = C2.name ";
                            sqlConsulta = sqlConsulta + " and C1.Agno = C2.Agno ";
                            sqlConsulta = sqlConsulta + " and C1.Mes = C2.Mes ";
                            sqlConsulta = sqlConsulta + " and C1.dia = C2.dia ";
                            sqlConsulta = sqlConsulta + " ORDER BY C1.name, C1.Agno, C1.Mes, C1.dia ";

                            conn.Open();
                            OleDbCommand comando = new OleDbCommand(sqlConsulta, conn);
                            OleDbDataReader lectura = comando.ExecuteReader();

                            entro = "NO";

                            while (lectura.Read())
                            {
                                entro = "SI";
                                cuenta = cuenta +1; ;
                                campo1 = lectura[0].ToString();
                                campo2 = lectura[1].ToString();
                                campo3 = lectura[2].ToString();
                                campo4 = lectura[3].ToString();

                                Entrada = System.Convert.ToDateTime(lectura[4].ToString());
                                Salida = System.Convert.ToDateTime(lectura[5].ToString());

                                estiloE = estiloS = "";
                                DIFERENCIA = System.Convert.ToDateTime(lectura[5].ToString()) - System.Convert.ToDateTime(lectura[4].ToString());

                                if (Entrada.DayOfWeek == DayOfWeek.Friday)
                                {
                                    fecha = System.Convert.ToDateTime(lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + " 08:01:00");

                                }
                                else
                                {
                                    fecha = System.Convert.ToDateTime(lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + " 07:31:00");
                                }

                                DiferenciaEntrada = Entrada - System.Convert.ToDateTime(fecha);
                                DiferenciaPrimeraSalida = Entrada - PrimeraSalida;

                                if (DiferenciaEntrada.TotalMinutes <= 0)
                                {
                                    estiloE = "style=\"color: #008000\"";
                                }
                                else if (DiferenciaEntrada.TotalMinutes <= 15)
                                {
                                    estiloE = "style = \"color: #FF9900\"";
                                }
                                else
                                {
                                    estiloE = "style=\"color: #FF0000\"";
                                }
                                ValorrReader4 = lectura[4].ToString();

                                fecha = System.Convert.ToDateTime(lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + " 17:00:00");

                                DiferenciaSalida = Salida - System.Convert.ToDateTime(fecha);

                                if (DiferenciaSalida.TotalMinutes < 0)
                                {
                                    estiloS = "style=\"color: #FF0000\"";

                                }
                                else if (DiferenciaSalida.TotalMinutes <= 60)
                                {
                                    estiloS = "style=\"color: #008000\"";

                                }
                                else
                                {
                                    estiloS = "style = \"color: #FF9900\"";
                                }
                                if (DIFERENCIA.TotalMinutes == 0)
                                {
                                    estiloS = "style=\"color: #090909\"";
                                    cuerpop = cuerpop + "<TR><TD>" + lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + "</TD>";
                                    cuerpop = cuerpop + "<TD " + estiloE + ">" + ValorrReader4.Substring(10, ValorrReader4.Length - 10) + "</TD>";
                                    cuerpop = cuerpop + "<TD " + estiloS + ">N.R.</TD>";
                                    cuerpop = cuerpop + "<TD>" + DIFERENCIA.ToString() + "</TD></TR>";
                                }
                                else
                                {
                                    cuerpop = cuerpop + "<TR><TD>" + lectura[2].ToString() + "/" + lectura[1].ToString() + "/" + lectura[0].ToString() + "</TD>";
                                    cuerpop = cuerpop + "<TD " + estiloE + ">" + ValorrReader4.Substring(10, ValorrReader4.Length - 10) + "</TD>";
                                    cuerpop = cuerpop + "<TD " + estiloS + ">" + Salida.ToString().Substring(10, Salida.ToString().Length - 10) + "</TD>";
                                    cuerpop = cuerpop + "<TD>" + DIFERENCIA.ToString() + "</TD></TR>";
                                }
                            }
                            if (entro == "NO")
                            {
                                if (DateTime.Now.AddDays(nDia - intHace).DayOfWeek != DayOfWeek.Saturday)
                                {
                                    if (DateTime.Now.AddDays(nDia - intHace).DayOfWeek != DayOfWeek.Sunday)
                                    {
                                        cuenta = cuenta +1;
                                        estiloE = estiloS = "style=\"color: #090909\"";
                                        cuerpop = cuerpop + "<TR><TD>" + diaini.ToString().PadLeft(2, '0') + "/" + mesini.ToString().PadLeft(2, '0') + "/" + agnoini.ToString() + "</TD>";
                                        cuerpop = cuerpop + "<TD " + estiloE + ">N.R.</TD>";
                                        cuerpop = cuerpop + "<TD " + estiloS + ">N.R.</TD>";
                                        cuerpop = cuerpop + "<TD>N.R.</TD></TR>";
                                    }
                                }
                            }
                            conn.Close();
                        }
                        cuerpop = "<TR><TD rowspan = " + cuenta.ToString() + ">" + nombrefuncionario + "</TD>" + cuerpop;
                        cuenta = 1;
                        cuerpo = cuerpo + cuerpop;
                        cuerpop = "";

                    }

                    cuerpo = cuerpo + cuerpop + "</table><br/><br/>";

                    cuerpo = cuerpo + " <table border = \"1\" > ";
                    cuerpo = cuerpo + " <tr> ";
                    cuerpo = cuerpo + "  <td colspan = 2 align = \"center\" style = \"color: #FFFFFF; background-color: #FF0000\" > NOMENCLATURA </ td > ";
                    cuerpo = cuerpo + " </tr> ";
                    cuerpo = cuerpo + " <tr> ";
                    cuerpo = cuerpo + " <td align = \"center\" style = \"color: #FFFFFF; background-color: #FF0000\" > Entrada </ td > ";
                    cuerpo = cuerpo + " <td align = \"center\" style = \"color: #FFFFFF; background-color: #FF0000\" > Salida </ td > ";
                    cuerpo = cuerpo + " </tr> ";
                    cuerpo = cuerpo + " <tr> ";
                    cuerpo = cuerpo + " <td><span style = \"color: #008000\" > &#9679;</span> Indica que el Funcionario entró antes o a la hora definida<br/> ";
                    cuerpo = cuerpo + " <span style = \"color: #FF9900\" > &#9679;</span>  Indica que el Funcionario entró hasta 15 minutos después de la hora definida<br/> ";
                    cuerpo = cuerpo + " <span style = \"color: #FF0000\" > &#9679;</span>  Indica que el Funcionario entró después de 15 minutos de la hora definida<br/></ td > ";
                    cuerpo = cuerpo + " <td ><span style = \"color: #008000\" > &#9679;</span> Indica que el Funcionario salió entre las 5:00 PM y 6:00 PM<br/> ";
                    cuerpo = cuerpo + " <span style = \"color: #FF9900\" > &#9679;</span>  Indica que el Funcionario salió después de las 6:00 PM<br/> ";
                    cuerpo = cuerpo + " <span style = \"color: #FF0000\" > &#9679;</span>  Indica que el funcionario salió antes de las 5:00PM<br/></td> ";
                    cuerpo = cuerpo + " </tr> ";
                    cuerpo = cuerpo + " <tr><td colspan='2'>N.R.: Indica que no hay Registro de Información<br/> </td></tr>";
                    cuerpo = cuerpo + " </table> ";


                    cuerpo = cuerpo + "<br><br><br>Cordialmente.</body>";
                    cuerpo = cuerpo + "</body>";





                    emaildestino = correoa;
                    CopiaOculta = "";
                    asunto = "[BIOMETRICO]Control de acceso";
                    adjunto1 = "";
                    adjunto2 = "";




                   enviarcorreo(emaildestino, CopiaOculta, cuerpo, asunto, adjunto1, adjunto2);
                    if (area == 5)
                    {
                        emaildestino = "jairo.hamon@serviciodeempleo.gov.co";
                        enviarcorreo(emaildestino, CopiaOculta, cuerpo, asunto, adjunto1, adjunto2);
                        emaildestino = "pedro.beltran@serviciodeempleo.gov.co";
                        enviarcorreo(emaildestino, CopiaOculta, cuerpo, asunto, adjunto1, adjunto2);
                        
                    }

                }


            }
            catch (Exception ex)
            {
                label2.Text = ex.Message;
                emaildestino = "jairo.hamon@serviciodeempleo.gov.co";
                enviarcorreo(emaildestino, CopiaOculta, ex.Message, "BIOMETRICOS: ERROR", adjunto1, adjunto2);
            }
            finally
            {
                conn.Close();
                this.Close();
            }



        }

    }
}
