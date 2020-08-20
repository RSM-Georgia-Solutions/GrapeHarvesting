using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace GrapeHarvestingExcelImport.Models
{
    public class OutgoingPaymentModel
    {
        public InvoiceModel Invoice { get; set; }

        public string Add()
        {
            Payments outgoingPayment = (Payments)DiManager.Company.GetBusinessObject(BoObjectTypes.oVendorPayments);
            outgoingPayment.CardCode = Invoice.CardCode;
            outgoingPayment.DocDate = Invoice.PostingDate;
            outgoingPayment.TransferSum = Invoice.Price * Invoice.Quantity;
            outgoingPayment.TransferAccount = "1430";
            outgoingPayment.BPLID = Invoice.BplId;
            var res = outgoingPayment.Add();
            if (res != 0)
            {
                var err = DiManager.Company.GetLastErrorDescription();
                string errorMessage = $"Error Outgoing Payment Bp - {Invoice.CardCode} : {err}";
                //Application.SBO_Application.MessageBox(errorMessage);
                return errorMessage;
            }

            var docEntry = DiManager.Company.GetNewObjectKey();
            Recordset recSet = (Recordset)DiManager.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery($"SELECT TransId, DocNum FROM OVPM WHERE DocEntry = {docEntry}");
            int transId = int.Parse(recSet.Fields.Item($"TransId").Value.ToString());
            int paymentDocNum = int.Parse(recSet.Fields.Item($"DocNum").Value.ToString());
            JournalEntries journalEntry = (JournalEntries)DiManager.Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
            journalEntry.GetByKey(transId);
            journalEntry.Lines.CostingCode2 = Invoice.CostCenter;
            var resje = journalEntry.Update();
            if (resje == 0) return DiManager.Company.GetNewObjectKey(); ;
            {
                var err = DiManager.Company.GetLastErrorDescription();
                string errorMessageJe = $"Error Update Entry Outgoing Payment DocNum - {paymentDocNum} : {err}";
                return errorMessageJe;
            }
        }
    }
}
