using BusinessPartnerSpecialPrice;
using System;

namespace BusinessPartnerSpecialPrice.Classes
{
    public class Connect
    {
        /// <summary>
        /// Method to connect DI API
        /// </summary>
        /// <returns></returns>
        public static void ConnectDI()
        {
            try
            {
                int setConnectionContextReturn = 0;

                string sCookie = null;
                string sConnectionContext = null;

                // First initialize the Company object
                Program.oCompany = new SAPbobsCOM.Company();

                // Acquire the connection context cookie from the DI API.
                sCookie = Program.oCompany.GetContextCookie();

                // Retrieve the connection context string from the UI API using the
                // acquired cookie.
                sConnectionContext = Program.oApplication.Company.GetConnectionContext(sCookie);

                // before setting the SBO Login Context make sure the company is not
                // connected

                if (Program.oCompany.Connected == true)
                {
                    Program.oCompany.Disconnect();
                }

                // Set the connection context information to the DI API.
                setConnectionContextReturn = Program.oCompany.SetSboLoginContext(sConnectionContext);

                Program.oCompany.Connect();

                if (setConnectionContextReturn != 0)
                {
                    Program.oApplication.StatusBar.SetText($"SBO DI Error {Program.oCompany.GetLastErrorDescription()}", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                Program.oApplication.StatusBar.SetText($"SBO DI has connected successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Program.oApplication.StatusBar.SetText($"SBO DI Error {ex.InnerException?.Message ??  ex.Message.ToString()}", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Method to connect to UI API
        /// </summary>
        public static void ConnectUI()
        {

            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

            try
            {
                // If there's no active application the connection will fail
                //SboGuiApi.AddonIdentifier = "56455230354241534953303030303030383639313A56313933383937363435301821FA2EA03BDD5BDED09A772C9FF0AE7D4EE6AF";
                SboGuiApi.Connect(sConnectionString);
                Program.oApplication = SboGuiApi.GetApplication(-1);
                Program.oApplication.StatusBar.SetText($"SBO UI has connected successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                //  Connection failed
                Program.oApplication.StatusBar.SetText($"SBO UI - ERROR: {ex.InnerException?.Message ?? ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }        
    }
}
