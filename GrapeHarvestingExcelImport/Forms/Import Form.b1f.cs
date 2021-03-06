
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using SAPbouiCOM.Framework;
using ExcelImportDll;
using GrapeHarvestingExcelImport.Controllers;
using GrapeHarvestingExcelImport.Forms;
using GrapeHarvestingExcelImport.Models;
using SAPbobsCOM;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;
using DataTable = System.Data.DataTable;

namespace GrapeHarvestingExcelImport
{

    [FormAttribute("GrapeHarvestingExcelImport.Import_Form", "Forms/Import Form.b1f")]
    class Import_Form : UserFormBase
    {
        public Import_Form()
        {
        }

        ExcelFileController excelFileController = new ExcelFileController();
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_6").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_5").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.PictureBox0 = ((SAPbouiCOM.PictureBox)(this.GetItem("Item_8").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_3").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_10").Specific));
            this.EditText1.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.EditText1_ChooseFromListBefore);
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_13").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.ComboBox3 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_15").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.VisibleAfter += new VisibleAfterHandler(this.Form_VisibleAfter);

        }

        private SAPbouiCOM.Button Button0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                ShowFolder showFolder = new ShowFolder(EditText0, ComboBox1);
                showFolder.Open();
            }
            catch (Exception exception)
            {
                Application.SBO_Application.MessageBox($"{exception.Message}. {exception.InnerException?.Message}. {exception.InnerException?.InnerException?.Message}");
            }

        }
        private void OnCustomInitialize()
        {
            Recordset recSet = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery($"SELECT BPLId,BPLName FROM OBPL");
            while (!recSet.EoF)
            {
                ComboBox0.ValidValues.Add(recSet.Fields.Item("BPLId").Value.ToString(), recSet.Fields.Item("BPLName").Value.ToString());
                recSet.MoveNext();
            }

            ComboBox0.Item.DisplayDesc = true;
            ComboBox3.ValidValues.Add("01", "Outgoing Payment");
            ComboBox3.ValidValues.Add("02", "A/P Invoice");
            Assembly entryAssembly = Assembly.GetEntryAssembly();
            if (entryAssembly == null) return;
            var path = System.IO.Path.GetDirectoryName(entryAssembly.Location) + "\\Media\\Sap.bmp";
            PictureBox0.Picture = path;
        }

        private SAPbouiCOM.Button Button1;

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (ComboBox0.Selected == null || string.IsNullOrWhiteSpace(EditText1.Value) || string.IsNullOrWhiteSpace(EditText0.Value) || ComboBox1.Selected == null || ComboBox2.Selected == null || ComboBox3.Selected == null)
            {
                Application.SBO_Application.SetStatusBarMessage("შეავსეთ ყველა ველი",
                    BoMessageTime.bmt_Short,
                    true);
                return;
            }

            try
            {
                Dictionary<int,string> errorRows = new Dictionary<int,string>();
                if (ComboBox3.Selected.Value == "02")//invoice
                {
                    DataTable data;
                    data = excelFileController.ReadExcelFile(ComboBox1.Selected.Description, EditText0.Value);

                    var inovices = GrapeController.ParseDataTableToInvoice(data);
                    int total = inovices.Count;
                    int increment = 0;
                    foreach (InvoiceModel invoiceModel in inovices)
                    {
                        invoiceModel.BplId = int.Parse(ComboBox0.Value);
                        invoiceModel.CostCenter = EditText1.Value;
                        invoiceModel.WareHouse = ComboBox2.Value;
                        var res = invoiceModel.Add();
                        int resInt;
                        bool isNumeric = int.TryParse(res, out resInt);
                        increment++;
                        if (!isNumeric)
                        {
                            //Application.SBO_Application.MessageBox($" Row - {increment} : {res} : {invoiceModel.PostingDate}");
                            errorRows.Add(increment, res);
                        }
                        Application.SBO_Application.SetStatusBarMessage($"{increment} Of {total}", BoMessageTime.bmt_Short, false);

                    }
                }

                if (ComboBox3.Selected.Value == "01")//payment
                {
                    var data = excelFileController.ReadExcelFile(ComboBox1.Selected.Description, EditText0.Value);
                    var inovices = GrapeController.ParseDataTableToInvoice(data);
                    int total = inovices.Count;
                    int increment = 0;
                    foreach (InvoiceModel invoiceModel in inovices)
                    {
                        invoiceModel.BplId = int.Parse(ComboBox0.Value);
                        invoiceModel.CostCenter = EditText1.Value;
                        invoiceModel.WareHouse = ComboBox2.Value;
                        OutgoingPaymentModel outgoingPayment = new OutgoingPaymentModel { Invoice = invoiceModel };
                        var res = outgoingPayment.Add();
                        int resInt;
                        bool isNumeric = int.TryParse(res, out resInt);
                        increment++;
                        if (!isNumeric)
                        {
                            errorRows.Add(increment,res);
                        }
                        Application.SBO_Application.SetStatusBarMessage($"{increment} Of {total}", BoMessageTime.bmt_Short, false);
                    }
                }
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"პროცესი დასრულებულია");
                if (errorRows.Count > 0)
                {
                    Application.SBO_Application.MessageBox($"შეცდომები მოხდა ხაზებზე :  {string.Join("- ", errorRows.Select(x=>x.Key +"-" + x.Value))}");
                }
            }
            catch (Exception e)
            {
                Application.SBO_Application.SetStatusBarMessage($"{e.Message}",
                    BoMessageTime.bmt_Short,
                    true);
            }
        }

        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.PictureBox PictureBox0;
        private ComboBox ComboBox0;
        private StaticText StaticText0;
        private EditText EditText1;
        private StaticText StaticText2;
        public string CostCenterCode { get; set; }

        private void EditText1_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = false;
            ListOfCostCenters centers = new ListOfCostCenters(this);
            centers.Show();

        }

        private Form _paramsForm;
        public void FillCostCenter()
        {
            _paramsForm.DataSources.UserDataSources.Item("UD_0").Value = CostCenterCode;
        }

        private void Form_VisibleAfter(SBOItemEventArg pVal)
        {
            _paramsForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
        }

        private StaticText StaticText3;
        private ComboBox ComboBox2;

        private void ComboBox0_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            var branch = ComboBox0.Selected.Value;
            Recordset recSet = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            while (ComboBox2.ValidValues.Count > 0)
            {
                ComboBox2.ValidValues.Remove(0, BoSearchKey.psk_Index);
            }
            recSet.DoQuery($"SELECT * FROM OWHS WHERE BPLId = '{branch}'");
            while (!recSet.EoF)
            {
                ComboBox2.ValidValues.Add(recSet.Fields.Item("WhsCode").Value.ToString(),
                    recSet.Fields.Item("WhsName").Value.ToString());
                recSet.MoveNext();
            }
        }

        private StaticText StaticText4;
        private ComboBox ComboBox3;
    }
}
