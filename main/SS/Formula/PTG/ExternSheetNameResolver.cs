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
                String wbName = externalSheet.WorkbookName;
                String sheetName = externalSheet.SheetName;
                if (wbName != null)
                {
                    sb = new StringBuilder(wbName.Length + sheetName.Length + cellRefText.Length + 4);
                    SheetNameFormatter.AppendFormat(sb, wbName, sheetName);
                }
                else
                {
                    sb = new StringBuilder(sheetName.Length + cellRefText.Length + 4);
                    SheetNameFormatter.AppendFormat(sb, sheetName);
                }
                if (externalSheet is ExternalSheetRange)
                {
                    ExternalSheetRange r = (ExternalSheetRange)externalSheet;
                    if (!r.FirstSheetName.Equals(r.LastSheetName))
                    {
                        sb.Append(':');
                        SheetNameFormatter.AppendFormat(sb, r.LastSheetName);
                    }
                }
            }
            else
            {
                String firstSheetName = book.GetSheetFirstNameByExternSheet(field_1_index_extern_sheet);
                String lastSheetName = book.GetSheetLastNameByExternSheet(field_1_index_extern_sheet);
                sb = new StringBuilder(firstSheetName.Length + cellRefText.Length + 4);
                if (firstSheetName.Length < 1)
                {
                    // What excel does if sheet has been deleted
                    sb.Append("#REF"); // note - '!' Added just once below
                }
                else
                {
                    SheetNameFormatter.AppendFormat(sb, firstSheetName);
                    if (!firstSheetName.Equals(lastSheetName))
                    {
                        sb.Append(':');
                        sb.Append(lastSheetName);
                    }
                }

            }
            sb.Append('!');
            sb.Append(cellRefText);
            return sb.ToString();
        }
    }
}
