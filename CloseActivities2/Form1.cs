using SAPFEWSELib;

using System;
using System.Windows.Forms;

namespace CloseActivities2
{
    public partial class Form1 : Form
    {
        public static string SAP_stat;
        public static int noerror = 0;

        public Form1()
        {
            InitializeComponent();
            SAPActive.OpenSap("SP1 - ECC 6.0 Production [PS_PM_SD_GRP]");
            SAPActive.Login("750", "T991059", "ee33ww22!@1", "EN");
        }

        public class SAPActive
        {
            public static GuiApplication SapGuiApp { get; set; }
            public static GuiConnection SapConnection { get; set; }
            public static GuiSession session { get; set; }
            public static object GlobalVariables { get; private set; }

            public static void OpenSap(string env)
            {
                SapGuiApp = new GuiApplication();

                string connectString = null;
                if (env.ToUpper().Equals("DEFAULT"))
                {
                    connectString = "SP1 - ECC 6.0 Production [PS_PM_SD_GRP]";
                }
                else
                {
                    connectString = env;
                }
                SapConnection = SapGuiApp.OpenConnection(connectString, Sync: true); //creates connection
                session = (GuiSession)SapConnection.Sessions.Item(0); //creates the Gui session off the connection you made 

                session.RecordFile = @"SampleScript.vbs";
                session.Record = true;
            }
            public static void Login(string myclient, string mylogin, string mypass, string mylang)
            {
                GuiTextField client = (GuiTextField)session.ActiveWindow.FindByName("RSYST-MANDT", "GuiTextField");
                GuiTextField login = (GuiTextField)session.ActiveWindow.FindByName("RSYST-BNAME", "GuiTextField");
                GuiTextField pass = (GuiTextField)session.ActiveWindow.FindByName("RSYST-BCODE", "GuiPasswordField");
                GuiTextField language = (GuiTextField)session.ActiveWindow.FindByName("RSYST-LANGU", "GuiTextField");

                client.SetFocus();
                client.Text = myclient;
                login.SetFocus();
                login.Text = mylogin;
                pass.SetFocus();
                pass.Text = mypass;
                language.SetFocus();
                language.Text = mylang;

                ClickButton(session);

                ProceedToCN25(session);

                session.Record = false;
            }
        }
        public static void ProceedToCN25(GuiSession session)
        {
            //GuiButton wnd1 = null;
            SAP_stat = string.Empty;
            string activityIn;

            //Start cn25
            session.StartTransaction("cn25");

            //Insert network number
            GuiCTextField network = (GuiCTextField)session.ActiveWindow.FindByName("CORUF-AUFNR", "GuiCTextField");
            network.SetFocus();
            network.Text = "2903039";

            //insert activity
            GuiTextField activity25 = (GuiTextField)session.ActiveWindow.FindByName("CORUF-VORNR", "GuiTextField");
            activity25.SetFocus();
            activity25.Text = "EEXP";
            activityIn = activity25.Text;

            ClickButton(session);

            int noerror = 0;
            try
            {
                GuiModalWindow wnd1 = (GuiModalWindow)session.FindById("wnd[1]");
                if (wnd1 != null)
                {
                    GuiModalWindow wnda = (GuiModalWindow)session.FindById("wnd[1]");
                    // string Session_value = session.FindById("wnd[1]").Text.Trim();

                    if (wnda.Text == "Status management: Confirm transaction" || wnda.Text == "Gestion des statuts : Confirmer opération")
                    {
                        // Press the button
                        GuiButton btnOpt1 = (GuiButton)session.FindById("wnd[1]/usr/btnOPTION1");
                        btnOpt1.SetFocus();
                        btnOpt1.Press();
                        noerror = 3;
                        goto ContinueExecution;
                    }
                    else if (wnda.Text == "Status management: Confirm order" || wnda.Text == "Gestion des statuts : Confirmer ordre")
                    {
                        GuiButton btnOpt1 = (GuiButton)session.FindById("wnd[1]/usr/btnOPTION1");
                        btnOpt1.SetFocus();
                        btnOpt1.Press();
                        noerror = 0;
                        goto ContinueExecution;
                    }
                    else if (wnda.Text == "Status management: Confirm order" || wnda.Text == "Status management: Confirm transaction")
                    {
                        GuiButton btnOpt1 = (GuiButton)session.FindById("wnd[1]/usr/btnOPTION1");
                        btnOpt1.SetFocus();
                        btnOpt1.Press();
                        noerror = 0;
                        goto ContinueExecution;
                    }
                }
                else
                {
                    goto confirmer_fermeture;
                }
            }
            catch
            {
                goto confirmer_fermeture;
            }

            try
            {
                // Get the status bar text
                GuiStatusbar statusBar = (GuiStatusbar)session.FindById("wnd[0]/sbar");
                string statusBarText = statusBar.Text;

                // Check if the status bar text contains "completed"
                if (statusBarText.Contains("completed"))
                {
                    noerror = 3;
                    goto ContinueExecution;
                }
            }
            catch (Exception ex)
            {
                var tatt = ex.Message; // Handle the exception (optional)
            }

            try
            {
                GuiModalWindow wnd0 = (GuiModalWindow)session.ActiveWindow.FindByName("AFRUD-LEKNW", "GuiModalWindow");
                // wnd0 = session.FindById("wnd[0]/usr/chkAFRUD-LEKNW");
                if (wnd0 != null)
                {
                    goto confirmer_fermeture;
                }
            }
            catch (Exception ex)
            {
                var tatta = ex.Message; // Handle the exception (optional)
            }



            confirmer_fermeture:



            try
            {
                // Select the checkbox
                GuiCheckBox chk1 = (GuiCheckBox)session.ActiveWindow.FindByName("AFRUD-LEKNW", "GuiCheckBox");
                GuiCheckBox chk2 = (GuiCheckBox)session.ActiveWindow.FindByName("AFRUD-AUERU", "GuiCheckBox");
                chk1.Selected = true;
                chk2.Selected = true;
                // Set focus to the checkbox
                chk2.SetFocus();

                // Press the button
                GuiButton btnOpt11 = (GuiButton)session.FindById("wnd[0]/tbar[0]/btn[11]");
                btnOpt11.SetFocus();
                btnOpt11.Press();

                // Check the text of the window
                GuiModalWindow wnd2 = (GuiModalWindow)session.FindById("wnd[2]");
                if (wnd2 != null)
                {
                    if (wnd2.Text == "Information " || wnd2.Text == "Warning")
                    {
                        noerror = 0;
                        // Press the button

                        // Session.FindById("wnd[2]/tbar[0]/btn[0]").press
                        GuiButton btnOpt12 = (GuiButton)session.FindById("wnd[2]/tbar[0]/btn[0]");
                        btnOpt12.SetFocus();
                        btnOpt12.Press();
                        btnOpt12.Press();

                        // Get the status text
                        GuiStatusbar SAP_stat = session.FindById("wnd[0]/sbar") as GuiStatusbar;
                        // SAP_stat = session.FindById("wnd[0]/sbar").Text;
                        noerror = 2;
                        goto ContinueExecution;
                    }
                }
                else
                {
                    GuiModalWindow wnd1b = (GuiModalWindow)session.FindById("wnd[1]");
                    if (wnd1b != null)
                    {
                        if (wnd1b.Text == "Information " || wnd1b.Text == "Warning")
                        {
                            GuiButton btnOpt12 = (GuiButton)session.FindById("wnd[1]/usr/btnOPTION1");
                            btnOpt12.SetFocus();
                            btnOpt12.Press();
                            btnOpt12.Press();

                            GuiButton btnOpt12b = (GuiButton)session.FindById("wnd[1]/tbar[0]/btn[0]");
                            btnOpt12b.SetFocus();
                            btnOpt12b.Press();
                            btnOpt12b.Press();

                            GuiStatusbar statusBar = (GuiStatusbar)session.FindById("wnd[0]/sbar");
                            // SAP_stat = session.FindById("wnd[0]/sbar").Text
                            noerror = 3;
                            goto ContinueExecution;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var tt32a = ex.Message; // Handle the exception (optional)
            }

            ContinueExecution:


            if (noerror < 1)
            {

            }
            else if (noerror == 1)
            {

            }
            else if (noerror == 2)
            {

            }
            else if (noerror == 3)
            {

            }

            else if (noerror == 4)
            {
            }

            else if (noerror == 5)
            {

            }

            else if (noerror == 6)
            {

            }

        }

        public static void ClickButton(GuiSession session)
        {
            //Press the green checkmark button 
            try
            {
                GuiButton btn = (GuiButton)session.FindById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[0]");
                btn.SetFocus();
                btn.Press();
            }
            catch (Exception ex)
            {

                var tttlbl = ex.Message;
            }
            //Press the back arrow 
            //GuiButton btn2 = (GuiButton)session.FindById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]");
            //btn2.SetFocus();
            //btn2.Press();
        }
    }
}

