using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GrapeHarvestingExcelImport.Models;
using SAPbobsCOM;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;
using DataTable = System.Data.DataTable;

namespace GrapeHarvestingExcelImport.Controllers
{
    public class GrapeController
    {
        public static List<InvoiceModel> ParseDataTableToInvoice(DataTable data)
        {
            List<InvoiceModel> invoices = new List<InvoiceModel>();
            List<DataRow> rows = data.AsEnumerable().ToList();
            object[] headersx = rows[0].ItemArray; //headers in actual excel
            Dictionary<string, int> excelIndexes = new Dictionary<string, int>();

            for (int i = 0; i < headersx.Length; i++)
            {
                string header = headersx[i].ToString(); //current header
                try
                {
                    excelIndexes.Add(header, i);
                }
                catch (Exception e)
                {
                    Application.SBO_Application.SetStatusBarMessage("დუბლირებული ველები Excel-ში",
                        BoMessageTime.bmt_Short, true);
                }
            }


            Dictionary<string, string> bpIdsAndCardCodes = new Dictionary<string, string>();
            Recordset recSet = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recSet2 = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet2.DoQuery($"select Series from NNM1 where Locked != 'Y' AND SeriesType = 'B' AND IsManual = 'N' AND DocSubType = 'S'");
            int series = int.Parse(recSet2.Fields.Item("Series").Value.ToString());

            recSet.DoQuery($"SELECT CardCode, isnull(VatIdUnCmp,LicTradNum) bpId FROM OCRD WHERE CardType = 'S' AND VatIdUnCmp is not null AND LicTradNum is not null");

            List<string> duplicates = new List<string>();

            int j = 0;
            while (!recSet.EoF)
            {

                string cardCode = recSet.Fields.Item("CardCode").Value.ToString();
                string id = recSet.Fields.Item("bpId").Value.ToString();

                if (id == string.Empty)
                {
                    recSet.MoveNext();
                    continue;
                }
                try
                {
                    bpIdsAndCardCodes.Add(id, cardCode);
                }
                catch (Exception ex)
                {
                    duplicates.Add(id);
                    recSet.MoveNext();

                }
                try
                {
                    recSet.MoveNext();
                    j++;
                }
                catch (Exception x)
                {
                    throw x;
                }
            }
            if (duplicates.Count > 0)
            {
                string ids = string.Join(Environment.NewLine, duplicates);
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"დუბლირებული პირადი ნომერი (ები) Sap - ში ID : {Environment.NewLine} {ids}");
            }

            foreach (DataRow row in rows.Skip(1))
            {
                string dateString = row[excelIndexes["თარიღი"]].ToString();
                double dateDouble;
                bool isNumeric = double.TryParse(dateString, out dateDouble);
                DateTime postingDate = isNumeric ? DateTime.FromOADate(dateDouble) : DateTime.Parse(dateString);

                string itemCode = row[excelIndexes["Item Code"]].ToString();
                string comment = row[excelIndexes["აქტი"]].ToString();
                double quantity = double.Parse(row[excelIndexes["კგ"]].ToString());
                string priceString = row[excelIndexes["ფასი"]].ToString();
                double price = double.Parse(priceString);
                string firsName = row[excelIndexes["სახელი"]].ToString();
                string lastName = row[excelIndexes["გვარი"]].ToString();
                string id = row[excelIndexes["პირადობის #"]].ToString();
                string adress = row[excelIndexes["მისამართი"]].ToString();
                string cardCode = string.Empty;

                InvoiceModel model = new InvoiceModel
                {
                    Comments = comment,
                    ItemCode = itemCode,
                    PostingDate = postingDate,
                    Price = price,
                    Quantity = quantity
                };                

               
                if (bpIdsAndCardCodes.ContainsKey(id))
                {
                    string cardCodeValue = bpIdsAndCardCodes[id].ToString();
                    Recordset recSet3 = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    var query = $@"SELECT DebPayAcct FROM OCRD WHERE CardCode = '{cardCodeValue}'";
                    recSet3.DoQuery(query);

                    if(recSet3.Fields.Item(0).Value.ToString() != "3112/001")
                        cardCode = CreateBP(bpIdsAndCardCodes, series, firsName, lastName, id, cardCode, model);
                    else 
                        model.CardCode = bpIdsAndCardCodes[id];
                }
                else
                {
                    cardCode = CreateBP(bpIdsAndCardCodes, series, firsName, lastName, id, cardCode, model);
                }
                invoices.Add(model);
            }

            return invoices;
        }

        private static string CreateBP(Dictionary<string, string> bpIdsAndCardCodes, int series, string firsName, string lastName, string id, string cardCode, InvoiceModel model)
        {
            BusinessPartners businessPartnerObject = (BusinessPartners)DiManager.Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);          
            businessPartnerObject.FederalTaxID = id;
            businessPartnerObject.UnifiedFederalTaxID = id;
            businessPartnerObject.CardName = firsName + ' ' + lastName;
            businessPartnerObject.CardType = BoCardTypes.cSupplier;
            businessPartnerObject.Series = series;
            businessPartnerObject.Territory = 1;
            businessPartnerObject.GroupCode = 104;
            businessPartnerObject.DebitorAccount = "3112/001";
            
            if(bpIdsAndCardCodes.ContainsKey(id))
                businessPartnerObject.UserFields.Fields.Item("U_ConnBpV").Value = bpIdsAndCardCodes[id].ToString();

            businessPartnerObject.AccountRecivablePayables.Add();
            int res = businessPartnerObject.Add();

            if (res == 0)
            {
                cardCode = DiManager.Company.GetNewObjectKey();
            }
            else
            {
                string err = DiManager.Company.GetLastErrorDescription();
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(err);
            }
            model.CardCode = cardCode;
            return cardCode;
        }
    }
}
