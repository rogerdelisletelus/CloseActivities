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
        DataTable bireTable;
        public string language;

        public Form1()
        {
            InitializeComponent();
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            RadioButton1.Checked = true; 
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
            if (emptyFields.Count > 0)
            {
                foreach (string field in emptyFields)
                {
                    label9.Text = (language == "1") ? "Cellule(s) vide(s) trouvée(s) dans Bire Excel :" + field : "Empty cell(s) found in Bire Excel: " + field;
                }
            }
            else
            {
                panel1.Visible = false;
                panel2.Visible = true;
                panel5.Visible = false;
                panel4.Visible = false;
                label11.Text = (language == "1") ? "Aucun champ vide trouvé." : "No empty fields found.";
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
                    catch (Exception ex)
                    {
                        label11.Text = "";
                        label10.Text = (ex.Message);
                        hasError = true;
                    }
                }
            }
            if (!hasError)
            {
                label11.Text = (language == "1") ? "Fichier Excel transféré avec succès vers la table SQL Bire" : "Excel file successfully transferred to SQL Bire table";
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = true;
                panel4.Visible = false;
                panel5.Visible = false;
            }
        }

        private void BtnOpenSAP_Click(object sender, EventArgs e)
        {
            txtPassword.PasswordChar = '*';
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = true;
            panel5.Visible = false;
            label11.Text = "";
            textBox7.Visible = false;
        }

        private void BtnOpenSapGui_Click(object sender, EventArgs e)
        {
            SAPActive.OpenSap("SP1 - ECC 6.0 Production [PS_PM_SD_GRP]");
            SAPActive.Login("750", "T991059", "ee33ww22!@1", "EN");
            SAPActive.SapSession.StartTransaction("cn22"); // ("VA01");

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

            public static void OpenSap(string env)
            {
                SAPActive.SapGuiApp = new GuiApplication();

                string connectString = null;
                if (env.ToUpper().Equals("DEFAULT"))
                {
                    connectString = "SP1 - ECC 6.0 Production [PS_PM_SD_GRP]";
                }
                else
                {
                    connectString = env;
                }
                SAPActive.SapConnection = SAPActive.SapGuiApp.OpenConnection(connectString, Sync: true); //creates connection
                SAPActive.SapSession = (GuiSession)SAPActive.SapConnection.Sessions.Item(0); //creates the Gui session off the connection you made
            }

            public static void Login(string myclient, string mylogin, string mypass, string mylang)
            {
                GuiTextField client = (GuiTextField)SAPActive.SapSession.ActiveWindow.FindByName("RSYST-MANDT", "GuiTextField");
                GuiTextField login = (GuiTextField)SAPActive.SapSession.ActiveWindow.FindByName("RSYST-BNAME", "GuiTextField");
                GuiTextField pass = (GuiTextField)SAPActive.SapSession.ActiveWindow.FindByName("RSYST-BCODE", "GuiPasswordField");
                GuiTextField language = (GuiTextField)SAPActive.SapSession.ActiveWindow.FindByName("RSYST-LANGU", "GuiTextField");

                client.SetFocus();
                client.Text = myclient;
                login.SetFocus();
                login.Text = mylogin;
                pass.SetFocus();
                pass.Text = mypass;
                language.SetFocus();
                language.Text = mylang;

                //Press the green checkmark button which is about the same as the enter key 
                GuiButton btn = (GuiButton)SapSession.FindById("/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[0]");
                btn.SetFocus();
                btn.Press();
            }
        }

        public void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton radioButton = sender as RadioButton;
            language = "1";
            BtnGetSQL.Text = "Choisir le Fichier Excel";
            BtnLoadSQL.Text = "Télécharger dans SQL Server";
            BtnOpenSAP.Text = "Ouvrir SAP GUI";
            label14.Text = "Entrez identification SAP";
            label12.Text = "Mot de Passe:";
            BtnOpenSapGui.Text = "Connecter";
            label1.Text = "Enregistrer #:";
            label2.Text = "Complété:";
            label3.Text = "Réseau:";
            label4.Text = "Activité:";
            label5.Text = "Statut:";
            label6.Text = "Début:";
            label7.Text = "Fin:";
            label13.Text = "SAP ID:";
            label12.Text = "Mot de Passe:";
            Button3.Text = "1- Executer confirmation réseau par activité";
            Button4.Text = "2-Executer  (TECO)";
            Button5.Text = "Remettre TECO (retour en arrière)";
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            language = "2";
            BtnGetSQL.Text = "Choose Excel File";
            BtnLoadSQL.Text = "Upload to SQL Server";
            BtnOpenSAP.Text = "Open SAP GUI";
            label14.Text = "Enter SAP Credential";
            label12.Text = "Password:";
            BtnOpenSapGui.Text = "Connect";
            label1.Text = "Record #:";
            label2.Text = "Completed:";
            label3.Text = "Network:";
            label4.Text = "Activity:";
            label5.Text = "Status:";
            label6.Text = "Start:";
            label7.Text = "End:";
            label13.Text = "SAP ID:";
            label12.Text = "Password:";
            Button3.Text = "1- Run network confirmation by activity";
            Button4.Text = "2-Execute (TECO)";
            Button5.Text = "Reset TECO (backtrack)";
        }
    }
}
