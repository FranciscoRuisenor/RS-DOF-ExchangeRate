using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using RS_LibMST;

namespace RS_DOF_ExchangeRate
{
    public partial class RS_Main : Form
    {
        public static bool WhatIsRebootAllowed = true;
        public static bool ShutDownInProgress = false;
        public static bool ForceApplicationReStart = false;
        public static int WhatIsHour2Start = 8;
        public static string WhatIsDomain = "@fortalezamf.mx";
        public static string WhatIsCustomerServiceMail = "Servicios.Clientes" + WhatIsDomain;
        public static string WhatIsHelpDeskMail = "Help.Desk" + WhatIsDomain;
        public static string WhatIsDateFormat = "dd/MM/yyyy";
        public static string WhatIsMyIP = "";
        public static string[] ODBCStringConn = new string[10];
        public static int nARR = 100;
        public static string[] RecDBGlobalParameters = new string[nARR];
        public static string[] RecDBConnectors = new string[nARR];

        public static string RSServer = "";
        enum zConnID { RS_DB, ADOBIN, MICROBOLMEX, FORTALEZA, DBWH, BANCOS, SIGMO_MySQL };
        enum zDBGlobalParameters
        {
            ID, sCompanyName, bEnabled, sDBVersion,
            sADOBIN, sMICROBOLMEX, sFORTALEZA, sDBWH,
            sMode, sNumU, sCuentaBANAMEXGtias, sCuentaBANCOMERXGtias,
            sCuentaBANAMEXPagos, sCuentaBANCOMERPagos,
            sPAR4Seguimiento, sBANCOS, sBOYFacuracionGlobal,
            sPapLegal, sHourInNormal, iTolerancia,
            sCuentaOXXPagos, sCuentaCOPPELPagos,
            sRECAShort, sRECALong, sRECAIndShort, sRECAIndLong, sMontoGarPrendarioConyuge, fIVA, RSAutoOFFVersion, sFecIniASERTA, sFecFinASERTA, bFecNacimientoBeneficiario,
            sSIGMO, sPapLegal2, sCuentaChedrauiPagos, sCuentaPayNetPagos
        };
        enum zDBConnector { ID, sMode, sConnID, sConnDesc, sServer, xStrConnect };

        public RS_Main()
        {
            InitializeComponent();

            // LOCAL VARIABLE INITIALIZATION
            RSServer = "192.168.100.156";
            // INIT Global Variables - Before General Initialization (OVERWRITE)
            RSGbl_Variable.APPGroup = "RS FaMF";           // Application group for the Registry
            RSGbl_Variable.OAPPCompany = "Fortaleza a Mi Futuro";           // Comapny Name
            RSGbl_Variable.OAPPName = "RS DOF ExchangeRate";      // Name of THIS Application, if not, it uses the RS Lib Name
            RSGbl_Variable.OAPPVersion = "1.0.0.5";             // This Application Version
            RSGbl_Variable.OAPPDescription = "Aplicación desarrollada y Licenciada para\n\nFortaleza a Mi Futuro\n\n";
            RSGbl_Variable.OAPPDescription += "Esta aplicación se encarga de actualizar el tipo de cambio para PLD";
            RSGbl_Variable.APPRSWAdvertising = true;
            // Production
            //DSN=RSFaMF;UID=RSFaMFUser;PWD=DBRSFaMFPwd;
            RSLib_ODBC.StrConnect = "C8FAE9E2E929E7059FCBDCEDEFEDD0F6EA98BBD1A1B9D5C7DB6D8289766F7674848578937F7882A9966D";
            RSGbl_Variable.RSVersionTable = "RSFaMFGlobalParameters";
            RSGbl_Variable.GlobalCompanyCode = "FaMF";
            RSGbl_Variable.RSUserMSTTable = "RSFaMFUserMST";
            RSGbl_Variable.sEmailLogo = "http://www.FortalezaMF.mx/image/zRSFaMFSmall.png";
            // CALL the RS_LIB_MST Initialization
            RSLib_Init.InitAssemblies();
            RSGbl_Variable.APPCopyR = "Copyright © Fortaleza a Mi Futuro (R)2012, 2017";
            RSLib_SendMail.SMTP_cc = WhatIsHelpDeskMail;
            RSLib_SendMail.SMTP_Emergency = WhatIsHelpDeskMail;

            RSGbl_Variable.IsNonSTOPApplication = true;
        }
        enum zMode { Init, Waiting2Start, InProcess, Completed, Error };
        int iMode = (int)zMode.Init;
        double dLastUSDValue = 0;
        DateTime xLastUpdate = new DateTime(2012, 1, 1);
        DateTime xNow = DateTime.Now;

        private void Form1_Load(object sender, EventArgs e)
        {
            setNow();
            lblDate.Visible = true;
            RegistryRead();
            WhatIsMyIP = ReadIPAddress();
            if (WhatIsMyIP.Length == 0)
            {
                this.Close();
                return;
            }
            //WhatIsMyPublicIP = ReadPublicIPAddress(1);
            // DB Connection
            ODBCStringConn[(int)zConnID.RS_DB] = RSLib_ODBC.StrConnect;
            RSGbl_Variable.CheckedDB = true;
            if (!RSLib_ODBC.RunLOG("Msg-000", "Start Application -> " + RSGbl_Variable.APPName + " Ver:" + RSGbl_Variable.APPVersion))
            {
                if (RSLib_ODBC.LastODBCError.IndexOf("No se encuentra el nombre del origen de datos y no se especificó ningún controlador predeterminado") != -1)
                {
                    // There is No ODBC Connector, lets create it.
                    if (RSLib_Registry.SetODBCForMSSQL(
                        RSLib_String.MyStringProvider(RSLib_ODBC.StrConnect, "DSN=", ";"),
                        RSServer,
                        RSLib_String.MyStringProvider(RSLib_ODBC.StrConnect, "UID=", ";")))
                    {
                        // Chack again if it worked
                        if (!RSLib_ODBC.RunLOG("Msg-000", "Start Application -> " + RSGbl_Variable.APPName + " Ver:" + RSGbl_Variable.APPVersion))
                        {
                            if (RSLib_ODBC.LastODBCError.IndexOf("No existe el servidor SQL Server o se ha denegado el acceso al mismo.") != -1)
                            {
                                DBDisconnected();
                                // UPS NO CONNECTION
                                MessageBox.Show("La Conexión con el Servidor SIMBANK no se pudo establecer!\n\nAsegure Conexión VPN y vuelva a intentar esta aplicación\n\n\nNo se puede continuar.",
                                    "Error de Conexión VPN", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                this.Close();
                                return;
                            }

                            DBDisconnected();
                            //ForceODBCConfig();
                            return;
                        }
                        // Continue... it is connected
                        MessageBox.Show("Conector ODBC creado con exito y conectado a la base de datos.\n\nLa aplicación se continuara ejecutando normalmente.", "Información");
                    }
                    else
                    {
                        DBDisconnected();
                        //ForceODBCConfig();
                        return;
                    }
                }
                if (RSLib_ODBC.LastODBCError.IndexOf("No existe el servidor SQL Server o se ha denegado el acceso al mismo.") != -1)
                {
                    DBDisconnected();
                    // UPS NO CONNECTION
                    MessageBox.Show("Conexión con el Servidor SIMBANK no se pudo establecer!\n\nAsegure Conexión VPN y vuelva a intentar esta aplicación\n\n\nNo se puede continuar.",
                        "Error de Conexión VPN", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                    return;
                }
            }
            this.Text = RSGbl_Variable.APPName + " Ver: " + RSGbl_Variable.APPVersion;
            if (!ReadGlobalParameters())
            {
                RSLib_CriticalError.CriticalErrorHandler(
                    "RSFaMFReport_Load", "Err-xxx",
                    "Fatal error, failure to read RSFaMFGlobalParameters",
                    "Thereie is an error reading Global Parameters",
                    "Error", "");
                ShutDown("No Global Parameters");
                return;
            }

            RSGbl_Variable.RootDirectory = RSGbl_Variable.DataPath + @"\Temp";

            ShowLastUpdateValue();

            timer1.Start();
        }
        private void ShowLastUpdateValue()
        {
            dLastUSDValue = 0;
            // Load latest USD Exchange Rate
            string xSQL = "declare @xMaxFecha Date=(select MAX(FECHA) Fecha from FORTALEZA_PRD.dbo.CF_VALOR_DOLAR); \n" +
                "select @xMaxFecha dMaxFecha, VALOR fUSD from FORTALEZA_PRD.dbo.CF_VALOR_DOLAR where FECHA=@xMaxFecha; \n";
            if (RSLib_ODBC.ReadSQL(xSQL, false))
            {
                if (!DateTime.TryParse(RSLib_ODBC.GlobalReturnValue[0], out xLastUpdate))
                    xLastUpdate = new DateTime(2012, 1, 1);
                dLastUSDValue = RSLib_String.AToF(RSLib_ODBC.GlobalReturnValue[1]);
                lblLastUpdate.Text = "$" + dLastUSDValue.ToString("#,##0.0000") + "  al " + xLastUpdate.ToString("dd/MM/yyyy");
                lblLastUpdate.Visible = true;
            }
            else
                RSLib_CriticalError.CriticalErrorHandler("Form Load on RS DOF ExchangeRate",
                    "Err-xxx", "Failure Reading last Update Rate",
                    "Falla leyendo el último Tipo de Cambio", "RS DOF ExhangeRate", xSQL);
        }
        private void setNow()
        {
            xNow = DateTime.Now;
            lblDate.Text = xNow.ToString("yyyy/MM/dd HH:mm:ss");
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();

            setNow();

            if (xNow.Second == 0 && xNow.Minute == 0 && xNow.Hour >= WhatIsHour2Start)
            {
                // Check if is today the lates, if does ignore
                if (xLastUpdate.Year != xNow.Year || xLastUpdate.Month != xNow.Month || xLastUpdate.Day != xNow.Day)
                {
                    //if (xNow.DayOfWeek == DayOfWeek.Saturday || xNow.DayOfWeek == DayOfWeek.Sunday)
                    //{
                    //    // Do Nothing
                    //}
                    //else
                    //{
                        // Let´s get the lates from DOF
                        Cursor xC = this.Cursor;
                        this.Cursor = Cursors.WaitCursor;

                        UpdateFromDOF();

                        this.Cursor = xC;
                    //}
                }
            }
            
            timer1.Start();
        }
        private void UpdateFromDOF()
        {
            string xProvider = "http://dof.gob.mx/indicadores.xml";
            WebRequest SendReq = HttpWebRequest.Create(xProvider);
            WebResponse GetRes = SendReq.GetResponse();

            System.IO.Stream StreamRes = GetRes.GetResponseStream();
            StreamReader ResStrmRdr = new StreamReader(StreamRes, Encoding.UTF8);

            string xDOFOriginal = ResStrmRdr.ReadToEnd();

            StreamRes.Close();
            GetRes.Close();

            string xDOF = RSLib_String.MyStringProvider(xDOFOriginal, "<title>DOLAR</title>", "</item>");
            xDOF = xDOF.Replace("\n", "");
            xDOF = xDOF.Replace("\t", "");
            xDOF = xDOF.Replace(" ", "");
            string xValue = RSLib_String.MyStringProvider(xDOF, "<description>", "</description>");
            string xDate = RSLib_String.MyStringProvider(xDOF, "<valueDate>", "</valueDate>");
            double fValue = RSLib_String.AToF(xValue);
            DateTime xLocalClock = DateTime.Now;
            string xSQL = "";
            if (!DateTime.TryParse(xDate, out xLocalClock) || fValue == 0)
            {
                xLocalClock = DateTime.Now;
                xSQL = "select distinct CONVERT(date,FECHA_FESTIVO) Festivo \n" +
                    "from FORTALEZA_PRD.dbo.CF_DIAS_FESTIVOS \n" +
                    "where FECHA_FESTIVO='" + xLocalClock.ToString(WhatIsDateFormat) + "' \n" +
                    "; \n";
                if (!RSLib_ODBC.ReadSQL(xSQL, false))
                {
                    if (xLocalClock.DayOfWeek == DayOfWeek.Saturday || xLocalClock.DayOfWeek == DayOfWeek.Sunday)
                    {
                        // IT is Saturday or Sunday - Nothing 2 save, then use the latest
                        fValue = dLastUSDValue;
                        xLocalClock = DateTime.Now;
                    }
                    else
                    {
                        // No Es feriado: Ignore
                        RSLib_CriticalError.CriticalErrorHandler("Form UpdateFromDOF on RS DOF ExchangeRate",
                            "Err-xxx", "Failure Getting Update Rate from DOF",
                            "Falla leyendo el WebService", "RS DOF ExhangeRate", xDOFOriginal);
                        return;
                    }
                }
                else
                {
                    // IT is feriado - Nothing 2 save, then use the latest
                    fValue = dLastUSDValue;
                    xLocalClock = DateTime.Now;
                }
            }
            xSQL = "insert into FORTALEZA_PRD.dbo.CF_VALOR_DOLAR (COD_EMPRESA,FECHA,VALOR) VALUES('00030','" + xLocalClock.ToString(WhatIsDateFormat) +
                "','" + fValue.ToString() + "'); \n";
            if (!RSLib_ODBC.RunSQL(xSQL))
            {
                RSLib_CriticalError.CriticalErrorHandler("Form UpdateFromDOF on RS DOF ExchangeRate",
                    "Err-xxx", "Failure Updating Table",
                    "Falla actualizando tabla", "RS DOF ExhangeRate", xSQL);
            }
            else
            {
                // Show
                ShowLastUpdateValue();
                // Send eMail Notificaciones
                string xMessage =
                    "Fecha: \t" + xLocalClock.ToString(WhatIsDateFormat) + "\n" +
                    "Valor Tipo de Cambio: \t" + fValue.ToString("#,##0.0000") + "\n";

                RSLib_ODBC.RunLOG("PLD-001", "Exchange Rate Updated -> " + xMessage.Replace("\n", " ").Replace("\t", "-"));

                string xBody = RSFormaTable(xMessage, "Tipo de Cambio - Sistema PLD", true, true, false, true, true);

                if (!RSLib_SendMail.SendMailMessage(RSLib_SendMail.SMTP_From, "Oficial.Cumplimiento" + WhatIsDomain, "", "Help.Desk" + WhatIsDomain, "[Tipo de Cambio] " + xLocalClock.ToString(WhatIsDateFormat), xBody, true))
                {
                    RSLib_CriticalError.CriticalErrorHandler("Form UpdateFromDOF on RS DOF ExchangeRate",
                        "Err-xxx", "Failure sending email Notificacion",
                        "Falla enviando correo de notificacion", "RS DOF ExhangeRate", "");
                }
            }
        }
        private void RS_Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            RegistryWrite();
        }
        private void RegistryRead()
        {
            string LastVersion = RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "Version", "0");
            if (LastVersion != RSGbl_Variable.APPVersion)
                RSLib_Registry.MySaveRegistry(RSGbl_Variable.APPName, "FirstTime", "0");
            RSLib_Registry.MySaveRegistry(RSGbl_Variable.APPName, "Version", RSGbl_Variable.APPVersion);
            int FirstTime = RSLib_String.AToI(RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "FirstTime", "0"));
            int x = RSLib_String.AToI(RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "X", this.Location.X.ToString()));
            int y = RSLib_String.AToI(RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "Y", this.Location.Y.ToString()));
            int w = RSLib_String.AToI(RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "W", this.Size.Width.ToString()));
            int h = RSLib_String.AToI(RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "H", this.Size.Height.ToString()));
            RSGbl_Variable.DataPath = Environment.GetEnvironmentVariable("APPDATA") + "\\RuisenorSW\\" + RSGbl_Variable.APPGroup;
            RSLib_File.CheckDirectoryIfNotExistCreateIt(RSGbl_Variable.DataPath);
            if (FirstTime == 0)
            {
                RSLib_Registry.MySaveRegistry(RSGbl_Variable.APPName, "ApplicationPath", RSGbl_Variable.ApplicationPath);
                RSLib_Registry.MySaveRegistry(RSGbl_Variable.APPName, "FirstTime", "1");
                if (LastVersion == "0")  // This is really the first time
                {
                    // DO Something
                }
            }
            RSGbl_Variable.ApplicationPath = RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "ApplicationPath", RSGbl_Variable.ApplicationPath);

            RSLib_ODBC.StrConnect = RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "ODBC Connect", RSLib_ODBC.StrConnect);
            RSLib_ODBC.StrConnect = RSLib_Encrypt.DesEncryptString(RSLib_ODBC.StrConnect);
            RSLib_ODBC.LevelLogs = Convert.ToInt32(RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "Level Logs", RSLib_ODBC.LevelLogs.ToString()));
            RSLib_SendMail.SMTP_Host = RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "SMTP Host", RSLib_SendMail.SMTP_Host);
            RSLib_SendMail.SMTP_Port = Convert.ToInt32(RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "SMTP Port", RSLib_SendMail.SMTP_Port.ToString()));
            RSLib_SendMail.SMTP_SSL = (
                RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "SMTP SSL", RSLib_SendMail.SMTP_SSL.ToString()) == "True"
                ? true : false);
            RSLib_SendMail.SMTP_USER = RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "SMTP USER", RSLib_SendMail.SMTP_USER);
            RSLib_SendMail.SMTP_PWD = RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "SMTP PWD", RSLib_SendMail.SMTP_PWD);
            RSLib_SendMail.SMTP_PWD = RSLib_Encrypt.DesEncryptString(RSLib_SendMail.SMTP_PWD);
            RSLib_SendMail.SMTP_From = RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "SMTP From", RSLib_SendMail.SMTP_From);

            //WhatIsCustomerServiceMail = RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "NEW Customer Service eMail", WhatIsCustomerServiceMail);

            RSGbl_Variable.LastUser = RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "LastUser", System.Environment.UserName);
            RSLib_ODBC.iSQLTimeOutGlobal = RSLib_String.AToI(RSLib_Registry.MyGetRegistry(RSGbl_Variable.APPName, "SQL Time Out", RSLib_ODBC.iSQLTimeOutGlobal.ToString()));
            if (RSLib_ODBC.iSQLTimeOutGlobal == -1)
                RSLib_ODBC.iSQLTimeOutGlobal = 90;

            // Copia LOGO Si no existe
            string SourcePath = RSGbl_Variable.ApplicationPath;
            string TargetPath = RSGbl_Variable.DataPath + @"\Image";
            RSLib_File.CheckDirectoryIfNotExistCreateIt(TargetPath);
            string xPNGFile = "zRSFaMFSmall";
            if (!File.Exists(Path.Combine(TargetPath, xPNGFile + ".png")))
                if (File.Exists(Path.Combine(SourcePath, xPNGFile + ".png")))
                    RSLib_File.MyBasicMoveOrCopyFile(
                        Path.Combine(SourcePath, xPNGFile + ".png"),
                        Path.Combine(TargetPath, xPNGFile + ".png"),
                        true
                        );

            CheckIfNoEmailSettingsThenOverride();

            RSGbl_Variable.GlobalUII = System.Environment.MachineName;
            this.Location = new Point(x, y);
            this.Size = new System.Drawing.Size(w, h);

        }
        void RegistryWrite()
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                RSLib_Registry.MySaveRegistry(RSGbl_Variable.APPName, "X", this.Location.X.ToString());
                RSLib_Registry.MySaveRegistry(RSGbl_Variable.APPName, "Y", this.Location.Y.ToString());
                RSLib_Registry.MySaveRegistry(RSGbl_Variable.APPName, "W", this.Size.Width.ToString());
                RSLib_Registry.MySaveRegistry(RSGbl_Variable.APPName, "H", this.Size.Height.ToString());
            }
        }
        private void CheckIfNoEmailSettingsThenOverride()
        {
            int nOverride = 0;
            if (RSLib_SendMail.SMTP_From == "YourGMail@gmail.com")
                nOverride++;
            if (RSLib_SendMail.SMTP_PWD.Length == 0)
                nOverride++;
            if (nOverride > 0)
            {
                RSLib_SendMail.SMTP_From = "RS.LibMST@gmail.com";
                RSLib_SendMail.SMTP_USER = RSLib_SendMail.SMTP_From;
                RSLib_SendMail.SMTP_Host = "smtp.gmail.com";
                RSLib_SendMail.SMTP_Port = 587;
                RSLib_SendMail.SMTP_SSL = true;

                RSLib_SendMail.SMTP_PWD = RSLib_Encrypt.DesEncryptString("CE0D0D0CC80FD7D5"); // InitsYears

                // Override NEW FORTALEZA email System SMTP
                RSLib_SendMail.SMTP_From = "Help.Desk@FortalezaMF.mx";
                RSLib_SendMail.SMTP_USER = RSLib_SendMail.SMTP_From;
                RSLib_SendMail.SMTP_Host = "smtp.office365.com";
                RSLib_SendMail.SMTP_Port = 587;
                RSLib_SendMail.SMTP_SSL = true;

                RSLib_SendMail.SMTP_PWD = RSLib_Encrypt.DesEncryptString("D7100E19FC43021783A6"); // Sistemas1!
                RSLib_SendMail.SMTP_PWD = RSLib_Encrypt.DesEncryptString("CE1C0E19C91C02F198A6"); // Just...2


                // Override NEW FORTALEZA email System POP
                RSLib_SendMail.POP_Host = "outlook.office365.com";
                RSLib_SendMail.POP_Port = 995;
                RSLib_SendMail.POP_SSL = true;

            }
        }
        private string ReadIPAddress()
        {

            string xIPs = "";
            string xNL = "";
            try
            {
                IPHostEntry xHosts = Dns.GetHostEntry("");
                foreach (IPAddress xHost in xHosts.AddressList)
                    if (xHost.AddressFamily == AddressFamily.InterNetwork)
                    {
                        xIPs += xNL + xHost.ToString();
                        xNL = " - ";
                    }
                if (xIPs == "127.0.0.1")
                {
                    xIPs = "";
                    MessageBox.Show("Error al tratar de localizar la IP de la computadora.\n\nAl parecer no se esta conectado a la RED!!!!\n\nFavor de verificar su conexión.", "Error FATAL",
                        MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error al tratar de localizar la IP de la computadora.\n\n" + e.Message, "Error FATAL",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            return xIPs;

        }
        private void DBDisconnected()
        {
            RSGbl_Variable.CheckedDB = false;
        }
        //private void ForceODBCConfig()
        //{
        //    RS_Config x = new RS_Config();
        //    x.ShowDialog();
        //    if (RSGbl_Variable.ApplicationReboot)
        //        ForceShutDown("Configuracion cambiada! Aplication va a reiniciar");
        //}
        public void ShutDown(string AnyCause)
        {
            if (ShutDownInProgress)
                return;
            //GlobalTable.Dispose();
            ShutDownInProgress = true;
            if (AnyCause.Length > 0)
                RSLib_ODBC.RunLOG("Err-xxx", AnyCause);
            Close();
        }
        public void ForceShutDown(string Reason)
        {
            if (WhatIsRebootAllowed)
            {
                System.Diagnostics.Process MyP = new System.Diagnostics.Process();
                MyP.StartInfo.FileName = RSGbl_Variable.ApplicationPath + "\\" + RSGbl_Variable.APPName + ".EXE";
                try
                {
                    //MessageBox.Show("Just Before LAUNCH Program:\n\n" + MyP.StartInfo.FileName);
                    MyP.Start();
                    //MessageBox.Show("Just After LAUNCH Program:\n\n" + MyP.StartInfo.FileName);
                }
                catch (Exception e)
                {
                    RSLib_ODBC.RunLOG("Err-999", "Reboot Failure: " + RSGbl_Variable.ApplicationPath + "\\" + RSGbl_Variable.APPName + ".EXE" + " Msg:" + e.Message + " Stack:" + e.StackTrace);
                    MessageBox.Show("Please start it again", "Reboot failure!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
            else
                MessageBox.Show("El reinicio de la aplicación no es posible.\n\nFavor de reiniciar la aplicación.",
                    RSGbl_Variable.APPName, MessageBoxButtons.OK, MessageBoxIcon.Information);

            // Start Application
            ShutDown(Reason);
        }
        private bool ReadGlobalParameters()
        {
            return
                GenericReadRecord("ReadGlobalParameters",
                    (int)zConnID.RS_DB,
                    "",
                    "RSFaMFGlobalParameters",
                    "bEnabled='True'",
                    RecDBGlobalParameters,
                    true
                    );
        }
        // DB Management
        /// <summary>
        /// Set the Right String connector for other DB than RSInterface, like DB_MASTER, etc. This will only be used one time in any Read or RUN ODBC Call
        /// </summary>
        /// <param name="iODBCID">enum parameter indicating what is the ODBC ID String [zConn]</param>
        /// 
        public static
        void UseODBCString(int iODBCID)
        {
            RSLib_ODBC.xStrConnect = ODBCStringConn[iODBCID];   // Use ALTERNATE STRING Connector
            RSLib_ODBC.IsxStrConnect = true;
            RSLib_ODBC.LastODBCError = "";
        }
        /// <summary>
        /// Generic Reading of Record using ODBC Connector. Provide ySQL or xTABLE+xWHERE parameters because it will take only one.
        /// </summary>
        /// <param name="xCaller">Name of the rutine that is calling</param>
        /// <param name="iODBCID">Index of the ODBC Array. Use -1 for RSInterlace</param>
        /// <param name="ySQL">Complete SQL sentece or as alternative use xTABLE + xWHERE</param>
        /// <param name="xTABLE">Table Name. If ySQL is passed this parameter is ignored</param>
        /// <param name="xWHERE">Elements for WHERE. If ySQL is passed this parameter is ignored</param>
        /// <param name="RecordReturn">This is the String[] that is passed back with the record read</param>
        /// <param name="bUserNotification">True if the user would receive a MessageBox in case of failure</param>
        /// <returns>True on Success</returns>
        public static
        bool GenericReadRecord(string xCaller, int iODBCID, string ySQL, string xTABLE,
            string xWHERE, string[] RecordReturn, bool bUserNotification)
        {
            bool RC = true;
            RSLib_String.MyStringArrayEmpty(RecordReturn);
            string xSQL = ySQL;
            if (ySQL.Length == 0)
            {
                xSQL = "SELECT * FROM " + xTABLE;
                if (xWHERE.Length > 0)
                    xSQL += " WHERE " + xWHERE;
            }
            if (iODBCID != -1)
                UseODBCString(iODBCID);
            else
                RSLib_ODBC.IsxStrConnect = false;   // Force Use the Deafult

            if (!RSLib_ODBC.ReadSQL(xSQL, bUserNotification))
            {
                if (bUserNotification)
                    RSLib_CriticalError.NonCriticalErrorHandler(xCaller,
                        "Err-xxx", "No Record found for Table: " + xTABLE + " WHERE " + xWHERE,
                        "Registro no encontrado!\n\nTABLA: " + xTABLE + "\n\nDONDE: " +
                        xWHERE + "\n\nFavor de verificar la Bitácora!",
                        "Error de lectura", xSQL);
                RC = false;
            }
            else
                RSLib_String.MyStringArrayCopy(RSLib_ODBC.GlobalReturnValue, RecordReturn, RSLib_ODBC.nGlobalReturnValue);
            return RC;
        }
        private string RSFormaTableHeader(bool bIsHeader, bool bFrame, string xAdditional)
        {
            string xRet = "";
            if (bIsHeader)
            {
                xRet = "<TABLE " + xAdditional;
                if (bFrame)
                    xRet += " BORDER=2 BGCOLOR=#ffffff CELLSPACING=0 border-color:#A3A3A3";
                xRet += ">";
                //xRet += "<FONT FACE=\"Calibri\" SIZE=3>";
                xRet += "<FONT SIZE=3>";
            }
            else
            {
                xRet += "</FONT>";
                xRet += "</TABLE>";
            }
            return xRet;
        }
        /// <summary>
        /// This routine generate an HTML Table based on StringCad (\t=split columns \n=split rows)
        /// </summary>
        /// <param name="sCad">String Chain that hols the data separated by \t and \n</param>
        /// <param name="xTitle">If bImage then Show the Title after</param>
        /// <param name="bFrame">True if Frame is required</param>
        /// <param name="bImage">True if an imgae shall be included</param>
        /// <param name="bComputerInfo">True if information like User, Computer and domain need to be included</param>
        /// <param name="bTimeStamp">True if Timestamp is required</param>
        /// <param name="bAppendRSWPowered">True if "Powered by RSW" at the end will be included</param>
        /// <returns></returns>
        private string RSFormaTable(string sCad, string xTitle, bool bFrame, bool bImage, bool bComputerInfo,
            bool bTimeStamp, bool bAppendRSWPowered)
        {
            string xBody = RSFormaTableHeader(true, bFrame, "");
            if (bImage)
            {
                string xImage = RSGbl_Variable.DataPath + @"\Image";
                RSLib_File.CheckDirectoryIfNotExistCreateIt(xImage);
                xImage = Path.Combine(xImage, "zRSFaMFSmall.png");
                if (File.Exists(xImage))
                {
                    string xNoSpacesImage = RSGbl_Variable.sEmailLogo;
                    xBody += "<TR><TD><a href=\"http://fortalezamf.mx\" target=\"_blank\"><img src=\"" + xNoSpacesImage + "\"></img></a></TD>";
                    if (xTitle.Length > 0)
                    {
                        //xBody += "<TD><P><FONT FACE=\"Comic Sans MS\"SIZE=5>" +  De acuerdo a COCO mejor usar un solo FONT
                        xBody += "<TD><P><FONT SIZE=5>" + xTitle + "</FONT></P></TD>\n";
                        //xBody += "<FONT FACE=\"Comic Sans MS\"SIZE=1>" +
                        //    RSLib_Browse.GetTimeStamp() + "</FONT></TD>\n";
                    }
                    xBody += "</TR>";
                }
            }
            string xTABLE = sCad;
            if (bComputerInfo)
            {
                xTABLE += "&nbsp;\t&nbsp;\n";
                xTABLE += "Usuario:\t" + System.Environment.UserName + "\n";
                xTABLE += "Computadora:\t" + System.Environment.MachineName + "\n";
                //xTABLE += "Dominio:\t" + System.Environment.UserDomainName + "\n";
                xTABLE += "IP:\t" + WhatIsMyIP + "\n";
                xTABLE += "RS Report Versión:\t" + RSGbl_Variable.APPVersion + "\n";
            }
            if (bTimeStamp)
                xTABLE += "Timestamp:\t" + RSLib_Browse.GetTimeStamp() + "\n";
            TextReader stringReader = new StringReader(xTABLE);
            while (true)
            {
                string sxLine = stringReader.ReadLine();
                if (sxLine == null || sxLine.Length == 0)
                    break;
                xBody += "<TR>";
                string[] xParts = sxLine.Split('\t');
                foreach (string xPart in xParts)
                    xBody += "<TD>" + xPart + "</TD>";
                xBody += "</TR>";
            }
            xBody += RSFormaTableHeader(false, bFrame, "");
            if (bAppendRSWPowered)
                xBody += "<Font Size=1>Powered by RSW</Font><br />";
            return xBody;
        }

        private void btnForceUpdate_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            System.Threading.Thread.Sleep(1000);
            Cursor xC = this.Cursor;
            this.Cursor = Cursors.WaitCursor;

            UpdateFromDOF();

            this.Cursor = xC;

            timer1.Start();
        }
    }
}
