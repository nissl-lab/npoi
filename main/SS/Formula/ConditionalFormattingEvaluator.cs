using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula
{
    public class ConditionalFormattingEvaluator
    {
        private WorkbookEvaluator workbookEvaluator;
        private IWorkbook workbook;

        private Dictionary<String, List<EvaluationConditionalFormatRule>> formats = new Dictionary<string, List<EvaluationConditionalFormatRule>>();
        private SortedDictionary<CellReference, List<EvaluationConditionalFormatRule>> values = new SortedDictionary<CellReference, List<EvaluationConditionalFormatRule>>();

        public ConditionalFormattingEvaluator(IWorkbook wb, IWorkbookEvaluatorProvider provider)
        {
            this.workbook = wb;
            this.workbookEvaluator = provider.GetWorkbookEvaluator();
        }
        protected WorkbookEvaluator WorkbookEvaluator
        {
            get
            {
                return workbookEvaluator;
            }
        }
        public List<EvaluationConditionalFormatRule> GetConditionalFormattingForCell(CellReference cellRef)
        {
            List<EvaluationConditionalFormatRule> rules = values[cellRef];

            if (rules == null)
            {
                // compute and cache them
                rules = new List<EvaluationConditionalFormatRule>();

                ISheet sheet;
                if (cellRef.SheetName != null)
                {
                    sheet = workbook.GetSheet(cellRef.SheetName);
                }
                else
                {
                    sheet = workbook.GetSheetAt(workbook.ActiveSheetIndex);
                }

                /*
                 * Per Excel help:
                 * https://support.office.com/en-us/article/Manage-conditional-formatting-rule-precedence-e09711a3-48df-4bcb-b82c-9d8b8b22463d#__toc269129417
                 * stopIfTrue is true for all rules from HSSF files, and an explicit value for XSSF files.
                 * thus the explicit ordering of the rule lists in #getFormattingRulesForSheet(Sheet)
                 */
                bool stopIfTrue = false;
                foreach (EvaluationConditionalFormatRule rule in GetRules(sheet))
                {

                    if (stopIfTrue)
                    {
                        continue; // a previous rule matched and wants no more evaluations
                    }

                    if (rule.Matches(cellRef))
                    {
                        rules.Add(rule);
                        stopIfTrue = rule.Rule.StopIfTrue;
                    }
                }
                rules.Sort();
                values.Add(cellRef, rules);
            }

            return rules;
        }
        public List<EvaluationConditionalFormatRule> GetConditionalFormattingForCell(ICell cell)
        {
            return GetConditionalFormattingForCell(GetRef(cell));
        }
        public static CellReference GetRef(ICell cell)
        {
            return new CellReference(cell.Sheet.SheetName, cell.RowIndex, cell.ColumnIndex, false, false);
        }

        protected List<EvaluationConditionalFormatRule> GetRules(ISheet sheet)
        {
            String sheetName = sheet.SheetName;
            List<EvaluationConditionalFormatRule> rules = formats[sheetName];
            if (rules == null)
            {
                if (formats.ContainsKey(sheetName))
                {
                    return new List<EvaluationConditionalFormatRule>();
                }
                ISheetConditionalFormatting scf = sheet.SheetConditionalFormatting;
                int count = scf.NumConditionalFormattings;
                rules = new List<EvaluationConditionalFormatRule>(count);
                formats.Add(sheetName, rules);
                for (int i = 0; i < count; i++)
                {
                    IConditionalFormatting f = scf.GetConditionalFormattingAt(i);
                    //optimization, as this may be expensive for lots of ranges
                    CellRangeAddress[] regions = f.GetFormattingRanges();
                    for (int r = 0; r < f.NumberOfRules; r++)
                    {
                        IConditionalFormattingRule rule = f.GetRule(r);
                        rules.Add(new EvaluationConditionalFormatRule(workbookEvaluator, sheet, f, i, rule, r, regions));
                    }
                }
                // need them in formatting and priority order so logic works right
                rules.Sort();
            }
            return rules;
        }
        public void ClearAllCachedFormats()
        {
            formats.Clear();
        }
        public void ClearAllCachedValues()
        {
            values.Clear();
        }
    }
}
