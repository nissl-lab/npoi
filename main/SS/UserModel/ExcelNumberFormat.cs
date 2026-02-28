using NPOI.SS.Formula;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.UserModel
{
    public class ExcelNumberFormat
    {
        private int idx;
        private String format;
        public ExcelNumberFormat(int idx, String format)
        {
            this.idx = idx;
            this.format = format;
        }
        public static ExcelNumberFormat From(ICellStyle style)
        {
            if (style == null) return null;
            return new ExcelNumberFormat(style.DataFormat, style.GetDataFormatString());
        }

        public static ExcelNumberFormat From(ICell cell, ConditionalFormattingEvaluator cfEvaluator)
        {
            if (cell == null) return null;

            ExcelNumberFormat nf = null;

            if (cfEvaluator != null)
            {
                // first one wins (priority order, per Excel help)
                List<EvaluationConditionalFormatRule> rules = cfEvaluator.GetConditionalFormattingForCell(cell);
                foreach (EvaluationConditionalFormatRule rule in rules)
                {
                    nf = rule.NumberFormat;
                    if (nf != null) break;
                }
            }
            if (nf == null)
            {
                ICellStyle style = cell.CellStyle;
                nf = ExcelNumberFormat.From(style);
            }
            return nf;
        }
        public int Idx
        {
            get { return idx; }
        }
        public String Format
        {
            get { return format; }
        }
    }
}
