using System;
using System.Text;

namespace NPOI.SS.Formula.PTG
{
    using NPOI.SS.Formula;
    public class ExternSheetNameResolver
    {

        public static String PrependSheetName(IFormulaRenderingWorkbook book, int field_1_index_extern_sheet, String cellRefText)
        {
            ExternalSheet externalSheet = book.GetExternalSheet(field_1_index_extern_sheet);
            StringBuilder sb;
            if (externalSheet != null)
            {
                String wbName = externalSheet.GetWorkbookName();
                String sheetName = externalSheet.GetSheetName();
                sb = new StringBuilder(wbName.Length + sheetName.Length + cellRefText.Length + 4);
                SheetNameFormatter.AppendFormat(sb, wbName, sheetName);
            }
            else
            {
                String sheetName = book.GetSheetNameByExternSheet(field_1_index_extern_sheet);
                sb = new StringBuilder(sheetName.Length + cellRefText.Length + 4);
                if (sheetName.Length < 1)
                {
                    // What excel does if sheet has been deleted
                    sb.Append("#REF"); // note - '!' added just once below
                }
                else
                {
                    SheetNameFormatter.AppendFormat(sb, sheetName);
                }
            }
            sb.Append('!');
            sb.Append(cellRefText);
            return sb.ToString();
        }
    }
}
