using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelImportDll
{

   public class ExcelFileController
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
                DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                sheets.AddRange(from DataRow drSheet in dt.Rows where drSheet["TABLE_NAME"].ToString().Contains("$") && !drSheet["TABLE_NAME"].ToString().Contains("FilterDatabase") select drSheet["TABLE_NAME"].ToString() into s select s.StartsWith("'") ? s.Substring(1, s.Length - 3) : s.Substring(0, s.Length - 1));
                connection.Close();
            }
            return sheets;
        }

        public string filePath { get; set; }

        public void ShowFolerBrowser()
        {

            Thread t = new Thread(() =>
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();


                DialogResult dr = openFileDialog.ShowDialog(new Form());
                if (dr == DialogResult.OK)
                {
                     filePath = openFileDialog.FileName;
                }
            }) {IsBackground = true}; // Kick off a new thread
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

    }
}
