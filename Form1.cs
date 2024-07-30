
using ExcelDataReader;

using Microsoft.Office.Interop.Excel;

using SAPFEWSELib;

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

using DataTable = System.Data.DataTable;

namespace CloseActivities
{
    public partial class Form1 : Form
    {
        public string filePath;
        public string connectionString;
        public SqlCommand SqlCmd;
        public SqlConnection connection;
        bool hasError = false;
        bool msg1 = false;
        bool msg2 = false;
        bool msg3 = false;
        bool msg4 = false;
        public DataTable bireTable;
        public string lang;
        public string gbfield;
        public static int noerror = 0;
        public static string SAP_stat;
        public Object gridview;
        public int sheetNum;

        public Form1()
        {
            InitializeComponent();

            RadioButton1.CheckedChanged += new EventHandler(RadioButtons_CheckedChanged);
            RadioButton2.CheckedChanged += new EventHandler(RadioButtons_CheckedChanged);
            RadioButton1.Checked = true;
            lang = "1";
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = true;
            panel5.Visible = false;
            panel7.Visible = false;
        }

        private void BtnGetSQL_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox7.Text = openFileDialog.FileName;
                filePath = openFileDialog.FileName;
                GetWorkSheet(filePath);
                // GetExcel(openFileDialog.FileName);

            }

            connectionString = ConfigurationManager.ConnectionStrings["BireDB"].ConnectionString;
        }

        public void GetWorkSheet(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            listBox1.Items.Clear();
            Workbook workbook = null;
            workbook = excelApp.Workbooks.Open(filePath);
            foreach (Worksheet sheet in workbook.Sheets)
            {
                listBox1.Items.Add(sheet.Name);
            }
            excelApp.Workbooks.Close();
            panel7.Visible = true;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            sheetNum = listBox1.SelectedIndex;
            bireTable = ReadExcelToDataTable(filePath, sheetNum);
            GetExcel(bireTable, filePath);
        }
        public void GetExcel(DataTable bireTable, string filePath)
        {
            // bireTable = ReadExcelToDataTable(filePath);

            List<string> emptyFields = GetEmptyFields(bireTable);
            label9.Text = "";
            label11.Text = "";
            msg1 = false;
            msg2 = false;
            if (emptyFields.Count > 0)
            {
                foreach (string field in emptyFields)
                {
                    msg2 = true;
                    gbfield = field;
                    label9.Text = (lang == "1") ? "Cellule(s) vide(s) trouvée(s) dans Bire Excel :" + field : "Empty cell(s) found in Bire Excel: " + field;
                }
            }
            else
            {
                msg1 = true;
                panel1.Visible = false;
                panel2.Visible = true;
                panel5.Visible = false;
                panel4.Visible = false;
                label11.Text = (lang == "1") ? "Aucune cellule(s) vide trouvé." : "No empty cell(s) found.";
            }
        }

        static DataTable ReadExcelToDataTable(string filePath, int sheetNum)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    DataTable dataTable = result.Tables[sheetNum];
                    return dataTable;
                }
            }
        }

        static List<string> GetEmptyFields(DataTable table)
        {
            List<string> emptyFields = new List<string>();

            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn column in table.Columns)
                {
                    if (column.ColumnName == "Project" || column.ColumnName == "Network" || column.ColumnName == "Activity")
                    {
                        if (IsFieldEmpty(row[column]))
                        {
                            emptyFields.Add($"Row {table.Rows.IndexOf(row) + 2}, Column {column.ColumnName}");
                        }
                    }
                }
            }

            return emptyFields;
        }
        static bool IsFieldEmpty(object field)
        {
            return field == null || string.IsNullOrWhiteSpace(field.ToString());
        }


        //public void BtnLoadSQL_Click(object sender, EventArgs e)
        //{
        //    msg1 = false;
        //    msg2 = false;
        //    msg3 = false;
        //    msg4 = false;
        //    if (bireTable != null)
        //    {
        //        using (connection = new SqlConnection(connectionString))
        //        {
        //            try
        //            {
        //                SqlCommand SqlCmd = new SqlCommand
        //                {
        //                    Connection = connection,
        //                    CommandType = CommandType.StoredProcedure,
        //                    CommandText = "Insert_into_dbo_BIRE_NILEC",
        //                };

        //                SqlCmd.Parameters.AddWithValue("@dataTable", bireTable);
        //                connection.Open();
        //                SqlCmd.ExecuteNonQuery();
        //            }
        //            catch (Exception)
        //            {
        //                msg4 = true;
        //                label11.Text = "";
        //                label10.Text = (lang == "1") ? "Erreur dans la procédure: Insert_into_dbo_BIRE_NILEC" : "Error in procedure: Insert_into_dbo_BIRE_NILEC";

        //                hasError = true;
        //            }
        //        }
        //    }
        //    if (!hasError)
        //    {
        //        msg3 = true;
        //        label11.Text = (lang == "1") ? "Fichier Excel transféré avec succès vers la table SQL Bire" : "Excel file successfully transferred to SQL Bire table";
        //        panel1.Visible = false;
        //        panel2.Visible = false;
        //        panel3.Visible = true;
        //        panel4.Visible = false;
        //        panel5.Visible = false;

        //    }
        //}

        private void BtnOpenSAP_Click(object sender, EventArgs e)
        {
            RadioButton1.Checked = true;
            txtPassword.PasswordChar = '*';
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = true;
            panel5.Visible = false;
            label11.Text = "";
            textBox7.Visible = false;
            msg3 = false;
        }

        private void BtnOpenSapGui_Click(object sender, EventArgs e)
        {
            SAPActive.OpenSap("SP1 - ECC 6.0 Production [PS_PM_SD_GRP]");
            SAPActive.Login("750", "T991059", "ee33ww22!@1", "EN");

            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = true;
            label11.Text = "";
            textBox7.Visible = false;
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
            GuiButton wnd = null;
            SAP_stat = string.Empty;
            string activityIn;
            //Start cn25
            session.StartTransaction("cn25");

            //GuiTextField commandField = (GuiTextField)session.FindById("wnd[0]/tbar[0]/okcd");
            //commandField.Text = "cn25";

            // Send the Enter key (VKey 0)
            //GuiModalWindow mainWindow = (GuiModalWindow)session.FindById("wnd[0]");
            //mainWindow.SendVKey(0);

            //Insert network number
            GuiCTextField network = (GuiCTextField)session.ActiveWindow.FindByName("CORUF-AUFNR", "GuiCTextField");
            network.SetFocus();
            network.Text = "2911985";

            //insert activity
            GuiTextField activity25 = (GuiTextField)session.ActiveWindow.FindByName("CORUF-VORNR", "GuiTextField");
            activity25.SetFocus();
            activity25.Text = "EEXP";
            activityIn = activity25.Text;

            ClickButton(session);

            //var wnd1 = session.FindById("wnd[0]/sbar").Type;
            //var wnd4 = session.FindById("wnd[0]/sbar").Id;
            //var wnd6 = session.FindById("wnd[0]/sbar").Name; 



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
                // Get the status bar text
                GuiStatusbar statusBar = (GuiStatusbar)session.FindById("wnd[0]/sbar");
                string statusBarText = statusBar.Text;

                // Print the status bar text to the console
                //MessageBox.Show(statusBarText);  
            }
            catch (Exception ex)
            {
                var tastt = ex.Message; // Handle the exception (optional)
                // This mimics the behavior of 'On Error Resume Next' by ignoring the error
            }

            try
            {
                // Find the window by ID
                wnd = (GuiButton)session.FindById("wnd[2]/tbar[0]/btn[43]");
            }
            catch (Exception ex)
            {
                var ttt = ex.Message; // Handle the exception (optional)
            }

            if (wnd != null)
            {
                try
                {
                    // Check the text of the window

                    GuiModalWindow wnd1 = (GuiModalWindow)session.FindById("wnd[1]");
                    if (wnd1.Text == "Status management: Confirm transaction" || wnd1.Text == "Gestion des statuts : Confirmer opération")
                    {
                        // Press the button
                        GuiButton btnOpt1 = (GuiButton)session.FindById("wnd[1]/usr/btnOPTION1");
                        btnOpt1.SetFocus();
                        btnOpt1.Press();
                        // ((GuiButton)session.FindById("wnd[1]/usr/btnOPTION1")).Press();

                        try
                        {
                            // Find the window by ID again
                            wnd = (GuiButton)session.FindById("wnd[2]/tbar[0]/btn[43]");
                        }
                        catch (Exception ex)
                        {
                            var ttt2 = ex.Message; // Handle the exception (optional)
                        }

                        if (wnd != null)
                        {
                            try
                            {
                                // Check if the text contains "Technical Information"
                                GuiButton btn = (GuiButton)session.FindById("wnd[2]/tbar[0]/btn[43]");
                                if (btn.Text.Contains("Technical Information"))
                                {
                                    // Press the buttons
                                    ((GuiButton)session.FindById("wnd[1]/usr/btnOPTION1")).Press();
                                    ((GuiButton)session.FindById("wnd[2]/tbar[0]/btn[12]")).Press();
                                    ((GuiButton)session.FindById("wnd[1]/usr/btnOPTION2")).Press();

                                    // Get the status text
                                    GuiStatusbar SAP_stat = session.FindById("wnd[0]/sbar") as GuiStatusbar;
                                    //SAP_stat = session.FindById("wnd[0]/sbar").Text;
                                    noerror = 4;
                                    goto ContinueExecution;
                                }
                            }
                            catch (Exception ex)
                            {
                                var ttt22 = ex.Message;  // Handle the exception (optional)
                            }
                        }

                        noerror = 0;
                        goto ContinueExecution;
                    }
                }
                catch (Exception ex)
                {
                    var ttt32 = ex.Message;  // Handle the exception (optional)
                }
            }

            try
            {
                // Check the text of the window
                GuiModalWindow wnd1 = (GuiModalWindow)session.FindById("wnd[1]");
                if (wnd1.Text == "Status management: Confirm order" || wnd1.Text == "Gestion des statuts : Confirmer ordre")
                {
                    // Press the button
                    GuiButton btnOpt1 = (GuiButton)session.FindById("wnd[1]/usr/btnOPTION1");
                    btnOpt1.SetFocus();
                    btnOpt1.Press();
                    //((GuiButton)session.FindById("wnd[1]/usr/btnOPTION1")).Press();
                    noerror = 0;
                    goto ContinueExecution;
                }
            }
            catch (Exception ex)
            {
                var ttt32a = ex.Message;  // Handle the exception (optional)
            }


            try
            {
                // Check the text of the window
                GuiModalWindow wnd1 = (GuiModalWindow)session.FindById("wnd[1]");
                if (wnd1.Text == "Warning")
                {
                    // Press the button
                    ((GuiButton)session.FindById("wnd[1]/usr/btnOPTION1")).Press();

                    // Get the status text
                    GuiStatusbar SAP_stat = session.FindById("wnd[0]/sbar") as GuiStatusbar;
                    //SAP_stat = session.FindById("wnd[0]/sbar").Text;
                    noerror = 2;
                    goto ContinueExecution;
                }
            }
            catch (Exception ex)
            {
                var ttat32a = ex.Message;// Handle the exception (optional)
            }


            try
            {
                // Select the checkbox
                GuiCheckBox chk = (GuiCheckBox)session.ActiveWindow.FindByName("AFRUD-AUERU", "GuiCheckBox");
                //GuiCheckbox chk = (GuiCheckbox)session.FindById("wnd[1]/usr/chkAFRUD-AUERU");
                chk.Selected = true;

                // Set focus to the checkbox
                chk.SetFocus();

                // Press the button
                ((GuiButton)session.FindById("wnd[1]/tbar[0]/btn[3]")).Press();

                // Check the text of the window
                GuiModalWindow wnd1 = (GuiModalWindow)session.FindById("wnd[1]");
                if (wnd1.Text == "Warning")
                {
                    // Press the button
                    ((GuiButton)session.FindById("wnd[1]/usr/btnOPTION1")).Press();

                    // Get the status text
                    GuiStatusbar SAP_stat = session.FindById("wnd[0]/sbar") as GuiStatusbar;
                    // SAP_stat = session.FindById("wnd[0]/sbar").Text;
                    noerror = 2;
                    goto ContinueExecution;
                }
            }
            catch (Exception ex)
            {
                var tt32a = ex.Message; // Handle the exception (optional)
            }


            try
            {
                // Check the text of the window
                GuiModalWindow wnd2b = (GuiModalWindow)session.FindById("wnd[2]");
                if (wnd2b.Text == "Back")
                {
                    // Press the button
                    ((GuiButton)session.FindById("wnd[2]/usr/btnSPOP-OPTION1")).Press();
                    noerror = 0;
                    goto ContinueExecution;
                }
            }
            catch (Exception ex)
            {
                var tt3z2a = ex.Message; // Handle the exception (optional)
            }

            try
            {
                // Get the text of the window and trim it
                GuiModalWindow wnd2c = (GuiModalWindow)session.FindById("wnd[2]");
                string wnd2Text = wnd2c.Text.Trim();

                // Check if the text is "Information" or "Warning"
                if (wnd2Text == "Information" || wnd2Text == "Warning")
                {
                    noerror = 0;

                    // Press the button twice
                    ((GuiButton)session.FindById("wnd[2]/tbar[0]/btn[0]")).Press();
                    ((GuiButton)session.FindById("wnd[2]/tbar[0]/btn[0]")).Press();
                }
            }
            catch (Exception ex)
            {
                var tt3za2a = ex.Message;// Handle the exception (optional)
            }

            try
            {
                // Get the text of the window and trim it
                GuiModalWindow wnd1 = (GuiModalWindow)session.FindById("wnd[1]");
                string wnd1Text = wnd1.Text.Trim();

                // Check if the text is "Warning" or "Information"
                if (wnd1Text == "Warning" || wnd1Text == "Information")
                {
                    // Press the button
                    ((GuiButton)session.FindById("wnd[1]/usr/btnOPTION1")).Press();

                    // Get the status text
                    GuiStatusbar SAP_stat = session.FindById("wnd[0]/sbar") as GuiStatusbar;
                    //SAP_stat = session.FindById("wnd[0]/sbar").Text;
                    noerror = 2;
                    goto ContinueExecution;
                }
            }
            catch (Exception ex)
            {
                var t3za2a = ex.Message; // Handle the exception (optional)
            }

        ContinueExecution:

            //if (noerror < 1)
            //{
            //    rs.Edit();
            //    //rs.status = session.FindById("wnd[0]/sbar").Text;
            //    GuiStatusbar SAP_stat = session.FindById("wnd[0]/sbar") as GuiStatusbar;

            //   // string ggg = SAP_stat.ToString();

            //    if (SAP_stat.ToString() == "S" || SAP_stat.ToString() == "")

            //   // if (session.FindById("wnd[0]/sbar").MessageType == "S" || session.FindById("wnd[0]/sbar").MessageType == "")
            //    {
            //        rs.Action = "go_close";
            //        rs.GO_ferme = "DONE";
            //        rs.SAP_status = SAP_stat;
            //        this.reseau = rs.Network;
            //        rs.Update();
            //    }
            //}




            ////Save the page
            //GuiButton btnSave = (GuiButton)session.FindById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[11]");
            //btnSave.SetFocus();
            //btnSave.Press();


            //GuiModalWindow ModalWindow1 = (GuiModalWindow)session.ActiveWindow.FindById("/app/con[0]/ses[0]/wnd[1]", "GuiModalWindow");
            //if (ModalWindow1.Text == "Status management: Confirm transaction" || ModalWindow1.Text == "Gestion des statuts : Confirmer opération")
            //{
            //    GuiButton btnOpt = (GuiButton)session.FindById("wnd[1]/usr/btnOPTION1");
            //    btnOpt.SetFocus();
            //    btnOpt.Press();
            //}





            //re-insert activity
            //GuiTextField activity22 = (GuiTextField)session.ActiveWindow.FindByName("CORUF-VORNR", "GuiTextField");
            //activity22.SetFocus();
            //activity22.Text = activityIn;

            //ClickButton(session);

            //Start cn22
            session.StartTransaction("cn22");

            //Press the green checkmark button
            //ClickButton(session);


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

        private void RadioButtons_CheckedChanged(object sender, EventArgs e)
        {
            BtnGetSQL.Text = (lang == "1") ? "Choisir le Fichier Excel" : "Choose Excel File";
            BtnLoadSQL.Text = (lang == "1") ? "Télécharger dans SQL Server" : "Upload to SQL Server";
            BtnOpenSAP.Text = (lang == "1") ? "Ouvrir SAP GUI" : "Open SAP GUI";
            label14.Text = (lang == "1") ? "Entrez identification SAP" : "Enter SAP Credential";
            label12.Text = (lang == "1") ? "Mot de Passe:" : "Password:";
            BtnOpenSapGui.Text = (lang == "1") ? "Connecter" : "Connect";
            label1.Text = (lang == "1") ? "Enregistrer #:" : "Record #:";
            label2.Text = (lang == "1") ? "Complété:" : "Completed:";
            label3.Text = (lang == "1") ? "Réseau:" : "Network:";
            label4.Text = (lang == "1") ? "Activité:" : "Activity:";
            label5.Text = (lang == "1") ? "Statut:" : "Status:";
            label6.Text = (lang == "1") ? "Début:" : "Start:";
            label7.Text = (lang == "1") ? "Fin:" : "End:";
            label8.Text = (lang == "1") ? "Choisir la feuille" : "Choose worksheet";
            label13.Text = (lang == "1") ? "SAP ID:" : "SAP ID:";

            Button3.Text = (lang == "1") ? "1- Executer confirmation réseau par activité" : "1- Run network confirmation by activity";
            Button4.Text = (lang == "1") ? "2-Executer  (TECO)" : "2-Execute (TECO)";
            Button5.Text = (lang == "1") ? "Remettre TECO ( retour en arrière )" : "Reset TECO ( backtrack )";

            if (msg1)
            {
                label11.Text = (lang == "1") ? "Aucune cellule(s) vide trouvé." : "No empty fields found.";
            }
            if (msg2)
            {
                label9.Text = (lang == "1") ? "Cellule(s) vide(s) trouvée(s) dans Bire Excel." : "Empty cell(s) found in Bire Excel. " + gbfield;
            }
            if (msg3)
            {
                label11.Text = (lang == "1") ? "Fichier Excel transféré avec succès vers la table SQL BIRE_NILEC" : "Excel file successfully transferred to SQL BIRE_NILEC table";
            }
            if (msg4)
            {
                label10.Text = (lang == "1") ? "Erreur dans la procédure: Insert_into_dbo_BIRE_NILEC" : "Error in procedure: Insert_into_dbo_BIRE_NILEC";
            }
        }
        public void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            lang = "1";
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            lang = "2";
        }

        static void Translate(string[] args)
        {
            // Initialize the SAP GUI scripting engine
            GuiApplication sapGuiApp = new GuiApplication();
            GuiConnection sapConnection = (GuiConnection)sapGuiApp.Connections.Item(0);
            GuiSession session = sapConnection.Children.Item(0) as GuiSession;

            // Find the status bar text
            GuiStatusbar statusBar = session.FindById("wnd[0]/sbar") as GuiStatusbar;
            string statusBarText = statusBar.Text;

            // Output the status bar text
            Console.WriteLine("Status Bar Text: " + statusBarText);
        }


    }
}



////using System;
////using SAPFEWSELib;

////namespace SAPScriptExample
////{
////    class Program
////    {
////        static void Main(string[] args)
////        {
////            // Initialize the SAP GUI scripting engine
////            GuiApplication sapGuiApp = new GuiApplication();
////            GuiConnection sapConnection = sapGuiApp.Connections.Item(0);
////            GuiSession session = sapConnection.Children.Item(0) as GuiSession;

////            // Find the status bar text
////            GuiStatusbar statusBar = session.FindById("wnd[0]/sbar") as GuiStatusbar;
////            string statusBarText = statusBar.Text;

////            // Output the status bar text
////            Console.WriteLine("Status Bar Text: " + statusBarText);
////        }
////    }
////}
