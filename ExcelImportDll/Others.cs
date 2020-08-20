using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelImportDll
{
    public static class Others
    {
        public static void SetSheetNames(IEnumerable<string> sheets, ComboBox combo)
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
}
