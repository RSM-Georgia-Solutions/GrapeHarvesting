using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace GrapeHarvestingExcelImport.Models
{
    public class InvoiceModel
    {
        public DateTime PostingDate { get; set; }
        public string ItemCode { get; set; }
        public string Comments { get; set; }
        public double Quantity { get; set; }
        public double Price { get; set; }
        public string CardCode { get; set; }
        public int BplId { get; set; }
        public int DocEntry { get; set; }
        public string CostCenter { get; set; }
        public string WareHouse { get; set; }

        public string Add()
        {
            Documents apInvoice = (Documents)DiManager.Company.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
            apInvoice.CardCode = CardCode;
            apInvoice.VatDate = PostingDate;
            apInvoice.DocDate = PostingDate;
            apInvoice.Comments = Comments;
            apInvoice.Lines.ItemCode = ItemCode;
            apInvoice.Lines.UnitPrice = Price;
            apInvoice.Lines.Quantity = Quantity;
            apInvoice.BPL_IDAssignedToInvoice = BplId;
            apInvoice.Lines.VatGroup = "VAT0";
            apInvoice.Lines.WarehouseCode = WareHouse;
            int res = apInvoice.Add();
            return res == 0 ? DiManager.Company.GetNewObjectKey() : DiManager.Company.GetLastErrorDescription();
        }
    }
}
