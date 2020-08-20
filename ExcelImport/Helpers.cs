using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAPbouiCOM;
using DataTable = System.Data.DataTable;

namespace ExcelImport
{
    public class ExcelImportDll
    {
        public DataTable ReadExcelFile(string sheetName, string path)
        {

            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                if (fileExtension.ToLower() == ".xls")
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";
                if (fileExtension.ToLower() == ".xlsx")
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1'";
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1'";

                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";

                    comm.Connection = conn;

                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }

                }
            }
        }

        public List<string> ToExcelsSheetList(string excelFilePath)
        {
            List<string> sheets = new List<string>();
            using (OleDbConnection connection =
                new OleDbConnection((excelFilePath.TrimEnd().ToLower().EndsWith("x"))
                    ? "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + excelFilePath + "';" + "Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1'"
                    : "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + excelFilePath + "';Extended Properties=Excel 8.0;"))
            {
                connection.Open();
                System.Data.DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                sheets.AddRange(from DataRow drSheet in dt.Rows where drSheet["TABLE_NAME"].ToString().Contains("$") && !drSheet["TABLE_NAME"].ToString().Contains("FilterDatabase") select drSheet["TABLE_NAME"].ToString() into s select s.StartsWith("'") ? s.Substring(1, s.Length - 3) : s.Substring(0, s.Length - 1));
                connection.Close();
            }
            return sheets;
        }
        public static void SetSheetNames(IEnumerable<string> sheets, SAPbouiCOM.ComboBox combo)
        {
            while (combo.ValidValues.Count > 0)
            {
                combo.ValidValues.Remove(combo.ValidValues.Count - 1, BoSearchKey.psk_Index);
            }

            foreach (var sheet in sheets)
            {
                combo.ValidValues.Add(sheet, "");
            }
            combo.Item.Enabled = true;
            combo.Select(0, BoSearchKey.psk_Index);
        }


    }

    public class ShowFolder
    {
        public string value { get; set; }
        public delegate void myDeleg(string value);
        public event myDeleg currFunc;

        //კონსტრუქტორი, მხოლოდ ინიციალიზაციისთვის. დეფოლთი კი ისეც აქვს მარა რავიცი რას აშავებს იყოს
        public ShowFolder()
        {

        }
        private void addPath1(string value, ExcelImportDll importController, EditText editText, SAPbouiCOM.ComboBox comboBox)
        {
            editText.Value = value;

            if (value != "")
            {
                var sheetNames = importController.ToExcelsSheetList(editText.Value);
                ExcelImportDll.SetSheetNames(sheetNames, comboBox);
            }
        }


        public void Load(string value, ExcelImportDll importController, EditText editText, SAPbouiCOM.ComboBox comboBox)
        {
            ShowFolder newFolder = new ShowFolder();
            //add function to event
            newFolder.currFunc += addPath1(value, importController, editText, comboBox);
            //run method of folder class
            newFolder.loadFolder();
        }
        //ცალკე სრედი მეთოდის გასახსნელად, რადგან ამ დროს აპლიკაციაა გაშვებული 1 ცალ სრედში და მასზე ფეილის გაშვება
        //ახალი სრედის გარეშე არ გამოდის
        public void loadFolder()
        {
            try
            {
                Thread ShowFolderBrowserThread = new System.Threading.Thread(setValue);

                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {

                    ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA);

                    ShowFolderBrowserThread.Start();

                }

                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {

                    ShowFolderBrowserThread.Start();

                    ShowFolderBrowserThread.Join();

                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message + " Thread Problem");
            }

        }

        //folder open method
        private void setValue()
        {
            try
            {
                NativeWindow nws = new NativeWindow();

                OpenFileDialog fdb = new OpenFileDialog();

                //  myProcs = Process.GetProcessesByName("SAP Business One");

                nws.AssignHandle(Process.GetProcessesByName("SAP Business One")[0].MainWindowHandle);

                // var mSboForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm; //GET ACTIVE FORM

                if (fdb.ShowDialog(nws) == DialogResult.OK)
                {
                    string test = fdb.FileName;
                    value = test;
                    currFunc(test);
                }
            }
            catch (Exception ex)
            {

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message + " Folder Cant Open", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
    }
}
