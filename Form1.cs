using ExcelDataReader;

using SAPFEWSELib;

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

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
        DataTable bireTable;
        public string lang;
        public string gbfield; 

        public Object gridview;

        public Form1()
        {
            InitializeComponent();

            RadioButton1.CheckedChanged += new EventHandler(RadioButtons_CheckedChanged);
            RadioButton2.CheckedChanged += new EventHandler(RadioButtons_CheckedChanged);
            RadioButton1.Checked = true;
            lang = "1";
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;


            panel5.Visible = false;
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
                GetExcel(openFileDialog.FileName);
                filePath = openFileDialog.FileName;
            }

            connectionString = ConfigurationManager.ConnectionStrings["BireDB"].ConnectionString;
        }

        public void GetExcel(string filePath)
        {
            bireTable = ReadExcelToDataTable(filePath);

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

        static DataTable ReadExcelToDataTable(string filePath)
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

                    DataTable dataTable = result.Tables[0];
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
                    if (IsFieldEmpty(row[column]))
                    {
                        emptyFields.Add($"Row {table.Rows.IndexOf(row) + 2}, Column {column.ColumnName}");
                    }
                }
            }

            return emptyFields;
        }
        static bool IsFieldEmpty(object field)
        {
            return field == null || string.IsNullOrWhiteSpace(field.ToString());
        }
        public void BtnLoadSQL_Click(object sender, EventArgs e)
        {
            msg1 = false;
            msg2 = false;
            msg3 = false;
            msg4 = false;
            if (bireTable != null)
            {
                using (connection = new SqlConnection(connectionString))
                {
                    try
                    {
                        SqlCommand SqlCmd = new SqlCommand
                        {
                            Connection = connection,
                            CommandType = CommandType.StoredProcedure,
                            CommandText = "Insert_into_dbo_BIRE_NILEC",
                        };

                        SqlCmd.Parameters.AddWithValue("@dataTable", bireTable);
                        connection.Open();
                        SqlCmd.ExecuteNonQuery();
                    }
                    catch (Exception)
                    {
                        msg4 = true;
                        label11.Text = "";
                        label10.Text = (lang == "1") ? "Erreur dans la procédure: Insert_into_dbo_BIRE_NILEC" : "Error in procedure: Insert_into_dbo_BIRE_NILEC";

                        hasError = true;
                    }
                }
            }
            if (!hasError)
            {
                msg3 = true;
                label11.Text = (lang == "1") ? "Fichier Excel transféré avec succès vers la table SQL Bire" : "Excel file successfully transferred to SQL Bire table";
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = true;
                panel4.Visible = false;
                panel5.Visible = false;

            }
        }

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
            public static GuiSession SapSession { get; set; }
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
                SapSession = (GuiSession)SapConnection.Sessions.Item(0); //creates the Gui session off the connection you made 

                SapSession.RecordFile = @"SampleScript.vbs";
                SapSession.Record = true;
            }

            public static void Login(string myclient, string mylogin, string mypass, string mylang)
            {
                GuiTextField client = (GuiTextField)SapSession.ActiveWindow.FindByName("RSYST-MANDT", "GuiTextField");
                GuiTextField login = (GuiTextField)SapSession.ActiveWindow.FindByName("RSYST-BNAME", "GuiTextField");
                GuiTextField pass = (GuiTextField)SapSession.ActiveWindow.FindByName("RSYST-BCODE", "GuiPasswordField");
                GuiTextField language = (GuiTextField)SapSession.ActiveWindow.FindByName("RSYST-LANGU", "GuiTextField");

                client.SetFocus();
                client.Text = myclient;
                login.SetFocus();
                login.Text = mylogin;
                pass.SetFocus();
                pass.Text = mypass;
                language.SetFocus();
                language.Text = mylang;

                ClickButton(SapSession);

                ProceedToCN22(SapSession);

                SapSession.Record = false;
            }
        }
        public static void ProceedToCN22(GuiSession SapSession)
        {
            string activityIn;
            //Start cn25
            SapSession.StartTransaction("cn25");

            //Insert network number
            GuiCTextField network = (GuiCTextField)SapSession.ActiveWindow.FindByName("CORUF-AUFNR", "GuiCTextField");
            network.SetFocus();
            network.Text = "2920606";

            //insert activity
            GuiTextField activity = (GuiTextField)SapSession.ActiveWindow.FindByName("CORUF-VORNR", "GuiTextField");
            activity.SetFocus();
            activity.Text = "CPRD";
            activityIn = "CPRD";

            //Press the green checkmark button
            ClickButton(SapSession);

            //Check the confirmation checkbox
            GuiCheckBox checkBox = (GuiCheckBox)SapSession.ActiveWindow.FindByName("AFRUD-AUERU", "GuiCheckBox");
            checkBox.Selected = true;
            checkBox.SetFocus();

            //Save the page
            GuiButton btnSave = (GuiButton)SapSession.FindById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[11]");
            btnSave.SetFocus();
            btnSave.Press();

            //re-insert activity
            GuiTextField activity22 = (GuiTextField)SapSession.ActiveWindow.FindByName("CORUF-VORNR", "GuiTextField");
            activity22.SetFocus();
            activity22.Text = activityIn; // "CPRD";

            //Start cn22
            SapSession.StartTransaction("cn22");

            //Press the green checkmark button
            ClickButton(SapSession); 
            
        }
        public static void ClickButton(GuiSession SapSession)
        {
            //Press the green checkmark button 
            GuiButton btn = (GuiButton)SapSession.FindById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[0]");
            btn.SetFocus();
            btn.Press();

            //Press the back arrow 
            //GuiButton btn2 = (GuiButton)SapSession.FindById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[3]");
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
    }
}
