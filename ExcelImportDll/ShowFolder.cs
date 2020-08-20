using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAPbouiCOM;
using Form = System.Windows.Forms.Form;

namespace ExcelImportDll
{
    public class ShowFolder
    {
        public string value { get; set; }
        public delegate void myDeleg(string value);
        public event myDeleg currFunc;
        public SAPbouiCOM.EditText ComEditText;
        public SAPbouiCOM.ComboBox ComboBox;




        //კონსტრუქტორი, მხოლოდ ინიციალიზაციისთვის. დეფოლთი კი ისეც აქვს მარა რავიცი რას აშავებს იყოს
        public ShowFolder(SAPbouiCOM.EditText editText, SAPbouiCOM.ComboBox comboBox = null)
        {
            ComEditText = editText;
            ComboBox = comboBox;
        }



        private void setPath(string value)
        {
            ComEditText.Value = value;

            if (ComboBox != null)
            {
                FillComboBox();
            }
        }




        public void Open()
        {
            this.loadFolder();
            this.currFunc += setPath;

        }

        private void FillComboBox()
        {
            while (ComboBox.ValidValues.Count > 0)
            {
                ComboBox.ValidValues.Remove(0, BoSearchKey.psk_Index);
            }

            ExcelFileController excelFileController = new ExcelFileController();
            var list = excelFileController.ToExcelsSheetList(value);
            for (int i = 0; i < list.Count; i++)
            {
                ComboBox.ValidValues.Add((i+1).ToString(), list[i]);
            }

            ComboBox.Item.DisplayDesc = true;

            ComboBox.Select(1, BoSearchKey.psk_Index);
        } 
        private void loadFolder()
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
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message + " Thread Problem", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }



        }



        //folder open method
        private void setValue()
        {
            try
            {
                Thread t = new Thread(() =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();




                    DialogResult dr = openFileDialog.ShowDialog(new Form());
                    if (dr == DialogResult.OK)
                    {
                        string fileName = openFileDialog.FileName;
                        value = fileName;
                        currFunc(value);
                    }
                });          // Kick off a new thread
                t.IsBackground = true;
                t.SetApartmentState(ApartmentState.STA);
                t.Start();





            }
            catch (Exception ex)
            {



                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message + " Folder Cant Open", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }



        }
    }
}
