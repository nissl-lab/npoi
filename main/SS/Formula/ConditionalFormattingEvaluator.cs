using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.SS.Formula
{
    /// <summary>
    /// <para>
    /// Evaluates Conditional Formatting constraints.
    /// </para>
    /// <para>
    /// For performance reasons, this class keeps a cache of all previously evaluated rules and cells.
    /// Be sure to call <see cref="clearAllCachedFormats()" /> if any conditional formats are modified, added, or deleted,
    /// and <see cref="clearAllCachedValues()" /> whenever cell values change.
    /// </para>
    /// </summary>
    public class ConditionalFormattingEvaluator
    {
        private WorkbookEvaluator workbookEvaluator;
        private IWorkbook workbook;
        /// <summary>
        /// <para>
        /// All the underlying structures, for both HSSF and XSSF, repeatedly go to the raw bytes/XML for the
        /// different pieces used in the ConditionalFormatting* structures.  That's highly inefficient,
        /// and can cause significant lag when checking formats for large workbooks.
        /// </para>
        /// <para>
        /// Instead we need a cached version that is discarded when definitions change.
        /// </para>
        /// <para>
        /// Sheets don't implement equals, and since its an interface,
        /// there's no guarantee instances won't be recreated on the fly by some implementation.
        /// So we use sheet name.
        /// </para>
        /// </summary>
        private Dictionary<String, List<EvaluationConditionalFormatRule>> formats = new Dictionary<string, List<EvaluationConditionalFormatRule>>();
        /// <summary>
    /// <para>
    /// Evaluating rules for cells in their region(s) is expensive, so we want to cache them,
    /// and empty/reevaluate the cache when values change.
    /// </para>
    /// <para>
    /// Rule lists are in priority order, as evaluated by Excel (smallest priority # for XSSF, definition order for HSSF)
    /// </para>
    /// <para>
    /// CellReference : equals().
    /// </para>
    /// </summary>
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

        /// <summary>
        /// <para>
        /// This checks all applicable <see cref="ConditionalFormattingRule"/>s for the cell's sheet,
        /// in defined "priority" order, returning the matches if any.  This is a property currently
        /// not exposed from <c>CTCfRule</c> in <c>XSSFConditionalFormattingRule</c>.
        /// </para>
        /// <para>
        /// Most cells will have zero or one applied rule, but it is possible to define multiple rules
        /// that apply at the same time to the same cell, thus the List result.
        /// </para>
        /// <para>
        /// Note that to properly apply conditional rules, care must be taken to offset the base
        /// formula by the relative position of the current cell, or the wrong value is checked.
        /// This is handled by <see cref="WorkbookEvaluator.Evaluate(String, CellReference, CellRangeAddressBase)" />.
        /// </para>
        /// </summary>
        /// <param name="cellRef">NOTE: if no sheet name is specified, this uses the workbook active sheet</param>
        /// <returns>Unmodifiable List of <see cref="EvaluationConditionalFormatRule"/>s that apply to the current cell value,
        /// in priority order, as evaluated by Excel (smallest priority # for XSSF, definition order for HSSF),
        /// or null if none apply
        /// </returns>
        public List<EvaluationConditionalFormatRule> GetConditionalFormattingForCell(CellReference cellRef)
        {
            List<EvaluationConditionalFormatRule> rules = values.TryGetValue(cellRef, out List<EvaluationConditionalFormatRule> value) ? value : null;
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

        /// <summary>
        /// <para>
        /// This checks all applicable <see cref="ConditionalFormattingRule"/>s for the cell's sheet,
        /// in defined "priority" order, returning the matches if any.  This is a property currently
        /// not exposed from <c>CTCfRule</c> in <c>XSSFConditionalFormattingRule</c>.
        /// </para>
        /// <para>
        /// Most cells will have zero or one applied rule, but it is possible to define multiple rules
        /// that apply at the same time to the same cell, thus the List result.
        /// </para>
        /// <para>
        /// Note that to properly apply conditional rules, care must be taken to offset the base
        /// formula by the relative position of the current cell, or the wrong value is checked.
        /// This is handled by <see cref="WorkbookEvaluator.Evaluate(String, CellReference, CellRangeAddressBase)" />.
        /// </para>
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>Unmodifiable List of <see cref="EvaluationConditionalFormatRule"/>s that apply to the current cell value,
        /// in priority order, as evaluated by Excel (smallest priority # for XSSF, definition order for HSSF),
        /// or null if none apply
        /// </returns>
        public List<EvaluationConditionalFormatRule> GetConditionalFormattingForCell(ICell cell)
        {
            return GetConditionalFormattingForCell(GetRef(cell));
        }
        public static CellReference GetRef(ICell cell)
        {
            return new CellReference(cell.Sheet.SheetName, cell.RowIndex, cell.ColumnIndex, false, false);
        }

        /// <summary>
        /// lazy load by sheet since reading can be expensive
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns>unmodifiable list of rules</returns>
        protected List<EvaluationConditionalFormatRule> GetRules(ISheet sheet)
        {
            String sheetName = sheet.SheetName;
            List<EvaluationConditionalFormatRule> rules = formats.TryGetValue(sheetName, out List<EvaluationConditionalFormatRule> value) ? value : null;
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
        /// <summary>
        /// Call this whenever rules are added, reordered, or removed, or a rule formula is changed
        /// (not the formula inputs but the formula expression itself)
        /// </summary>
        public void ClearAllCachedFormats()
        {
            formats.Clear();
        }
        /// <summary>
        /// <para>
        /// Call this whenever cell values change in the workbook, so condional formats are re-evaluated
        /// for all cells.
        /// </para>
        /// <para>
        /// TODO: eventually this should work like <see cref="EvaluationCache.notifyUpdateCell(int, int, EvaluationCell)" />
        /// and only clear values that need recalculation based on the formula dependency tree.
        /// </para>
        /// </summary>
        public void ClearAllCachedValues()
        {
            values.Clear();
        }

        /// <summary>
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns>unmodifiable list of all Conditional format rules for the given sheet, if any</returns>
        public List<EvaluationConditionalFormatRule> GetFormatRulesForSheet(string sheetName) {
            return GetFormatRulesForSheet(workbook.GetSheet(sheetName));
        }
    
        /// <summary>
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns>unmodifiable list of all Conditional format rules for the given sheet, if any</returns>
        public List<EvaluationConditionalFormatRule> GetFormatRulesForSheet(ISheet sheet) {
            return GetRules(sheet);
        }
    
        /// <summary>
        /// <para>
        /// Conditional formatting rules can apply only to cells in the sheet to which they are attached.
        /// The POI data model does not have a back-reference to the owning sheet, so it must be passed in separately.
        /// </para>
        /// <para>
        /// We could overload this with convenience methods taking a sheet name and sheet index as well.
        /// </para>
        /// </summary>
        /// <param name="sheet">containing the rule</param>
        /// <param name="conditionalFormattingIndex">of the <see cref="ConditionalFormatting"/> instance in the sheet's array</param>
        /// <param name="ruleIndex">of the <see cref="ConditionalFormattingRule"/> instance within the <see cref="ConditionalFormatting"/></param>
        /// <returns>unmodifiable List of all cells in the rule's region matching the rule's condition</returns>
        public List<ICell> GetMatchingCells(ISheet sheet, int conditionalFormattingIndex, int ruleIndex) {
            foreach (EvaluationConditionalFormatRule rule in GetRules(sheet)) {
                if (rule.Sheet.Equals(sheet) && rule.FormattingIndex == conditionalFormattingIndex && rule.RuleIndex == ruleIndex) {
                    return GetMatchingCells(rule);
                }
            }
            return [];
        }
    
        /// <summary>
        /// </summary>
        /// <param name="rule"></param>
        /// <returns>unmodifiable List of all cells in the rule's region matching the rule's condition</returns>
        public List<ICell> GetMatchingCells(EvaluationConditionalFormatRule rule) {
            List<ICell> cells = new List<ICell>();
            ISheet sheet = rule.Sheet;
        
            foreach (CellRangeAddress region in rule.Regions) {
                for (int r = region.FirstRow; r <= region.LastRow; r++) {
                    IRow row = sheet.GetRow(r);
                    if (row == null) {
                        continue; // no cells to check
                    }
                    for (int c = region.FirstColumn; c <= region.LastColumn; c++) {
                        ICell cell = row.GetCell(c);
                        if (cell == null) {
                            continue;
                        }
                    
                        List<EvaluationConditionalFormatRule> cellRules = GetConditionalFormattingForCell(cell);
                        if (cellRules.Contains(rule)) {
                            cells.Add(cell);
                        }
                    }
                }
            }
            return cells;
        }
    }
}
