using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using SAPFEWSELib;

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace CloseActivities
{
    public partial class Form1 : Form
    {
        string filePath;
        string connectionString;
        public SqlCommand SqlCmd;
        public SqlConnection connection;
        bool hasError = false;
        public Form1()
        {
            InitializeComponent();
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox7.Text = openFileDialog.FileName;
                ReadCells(openFileDialog.FileName);
                filePath = openFileDialog.FileName;
            }

            connectionString = ConfigurationManager.ConnectionStrings["BireDB"].ConnectionString;
        }

        public void ReadCells(string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                bool shouldBreak = false;
                foreach (Row row in sheetData.Elements<Row>())
                {
                    List<string> cellValues = new List<string>();
                    int cellIndex = 0;

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        // Get the column index of the cell
                        int columnIndex = GetColumnIndexFromName(GetColumnName(cell.CellReference));

                        // Add empty values for missing cells
                        while (cellIndex < columnIndex)
                        {
                            cellValues.Add(string.Empty);
                            cellIndex++;
                        }

                        // Add the cell value
                        cellValues.Add(GetCellValue(document, cell));
                        cellIndex++;
                    }

                    // Handle any trailing empty cells
                    while (cellIndex < row.Elements<Cell>().Count())
                    {
                        cellValues.Add(string.Empty);
                        cellIndex++;
                    }

                    // Process the cell values as needed
                    foreach (var value in cellValues)
                    {
                        panel2.Visible = true;
                        panel5.Visible = true;
                        if (value == "" || value == " ")
                        {
                            //Label label1 = new Label();
                            label9.Text = ("Empty cell(s) found in Bire Excel file at row: " + row.RowIndex);
                            panel2.Visible = false;
                            panel5.Visible = false;
                            shouldBreak = true;
                            break;
                        }
                    }
                    if (shouldBreak)
                    {
                        break;
                    }
                }

                if (!shouldBreak)
                {
                    panel1.Visible = false;
                    panel2.Visible = true;
                    panel5.Visible = false;
                    panel4.Visible = false;
                    label11.Text = "Excel Bire contains no blank field";
                }
            }
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue?.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            return value;
        }

        private static string GetColumnName(string cellReference)
        {
            string columnName = string.Empty;
            foreach (char ch in cellReference)
            {
                if (Char.IsLetter(ch))
                {
                    columnName += ch;
                }
                else
                {
                    break;
                }
            }
            return columnName;
        }
        private static int GetColumnIndexFromName(string columnName)
        {
            int columnIndex = 0;
            int factor = 1;

            for (int pos = columnName.Length - 1; pos >= 0; pos--)
            {
                columnIndex += factor * ((columnName[pos] - 'A') + 1);
                factor *= 26;
            }
            return columnIndex - 1;
        }

        public void Button2_Click(object sender, EventArgs e)
        {
            DataTable bireTable = ReadExcelFile(filePath);
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
                label11.Text = "Excel file successfully transferred to SQL Bire table";
                panel1.Visible = false;
                panel2.Visible = false;
                panel3.Visible = true;
                panel4.Visible = false;
                panel5.Visible = false; 
            }
        }

        static DataTable ReadExcelFile(string filePath)
        {
            DataTable bireTable = new DataTable();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();
                if (sheet != null)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                    if (sheetData != null)
                    {
                        bool isFirstRow = true;
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            DataRow dataRow = bireTable.NewRow();
                            int columnIndex = 0;

                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                string cellValue = GetCellValue(document, cell);

                                if (isFirstRow)
                                {
                                    bireTable.Columns.Add(cellValue);
                                }
                                else
                                {
                                    dataRow[columnIndex] = cellValue;
                                }
                                columnIndex++;
                            }

                            if (!isFirstRow)
                            {
                                bireTable.Rows.Add(dataRow);
                            }

                            isFirstRow = false;
                        }
                    }
                }
            }

            return bireTable;
        }
        public void Update_Info_Activitees()
        {
            using (connection = new SqlConnection(connectionString))
            {
                try
                {
                    SqlCommand SqlCmd = new SqlCommand
                    {
                        Connection = connection,
                        CommandType = CommandType.StoredProcedure,
                        CommandText = "Update_Info_Activitees",
                    };
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

        private void BtnOpenSAP_Click(object sender, EventArgs e)
        {
            SAPActive.OpenSap("SP1 - ECC 6.0 Production [PS_PM_SD_GRP]");
            SAPActive.Login("750", "T991059", "ee33ww22!@1", "EN");
            SAPActive.SapSession.StartTransaction("VA03");
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
                    //connectString = "1.0 Test ERP (DEFAULT)";
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
    }
}
