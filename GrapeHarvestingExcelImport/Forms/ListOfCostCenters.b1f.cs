using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using BoOrderType = SAPbouiCOM.BoOrderType;

namespace GrapeHarvestingExcelImport.Forms
{
    [FormAttribute("GrapeHarvestingExcelImport.Forms.ListOfCostCenters", "Forms/ListOfCostCenters.b1f")]
    class ListOfCostCenters : UserFormBase
    {
        private readonly Import_Form _importForm;
        public ListOfCostCenters(Import_Form importForm)
        {
            _importForm = importForm;
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_0").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Grid Grid0;

        private void OnCustomInitialize()
        {
            Grid0.DataTable.ExecuteQuery($"SELECT PrcCode as [კოდი],PrcName as [დასახელება] FROM OPRC WHERE DimCode = 2");
        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText0;

        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Grid0.Rows.SelectedRows.Clear();
            if (pVal.Row == -1)
            {
                return;
            }
            Grid0.Rows.SelectedRows.Add(pVal.Row);
        }

        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.Row == -1)
            {
                return;
            }
            string costCenterCode = Grid0.DataTable.GetValue("კოდი", Grid0.GetDataTableRowIndex(pVal.Row)).ToString();
            _importForm.CostCenterCode = costCenterCode;
            _importForm.FillCostCenter();
            Application.SBO_Application.Forms.ActiveForm.Close();
        }

        private void EditText0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Grid0.DataTable.ExecuteQuery($"SELECT PrcCode as [კოდი],PrcName as [დასახელება] FROM OPRC WHERE DimCode = 2 AND PrcName like N'%{EditText0.Value}%'");
        }

        private SAPbouiCOM.Button Button0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Grid0.Rows.SelectedRows.Count <= 0)
            {
                return;
            }

            int x = Grid0.Rows.SelectedRows.Item(0,
                BoOrderType.ot_SelectionOrder);
            string costCenterCode = Grid0.DataTable.GetValue("კოდი", Grid0.GetDataTableRowIndex(x)).ToString();
            _importForm.CostCenterCode = costCenterCode;
            _importForm.FillCostCenter();
            Application.SBO_Application.Forms.ActiveForm.Close();
        }
    }
}
