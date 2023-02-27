using EnumsNET;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula
{
    public enum OperatorEnum:byte
    {
        NO_COMPARISON=0,
        BETWEEN=1,
        NOT_BETWEEN=2,
        EQUAL=3,
        NOT_EQUAL=4,
        GREATER_THAN=5,
        LESS_THAN=6,
        GREATER_OR_EQUAL=7,
        LESS_OR_EQUAL=8
    }
    public class OperatorEnumHelper
    {
        public static bool IsValid(OperatorEnum @operator, object cellValue, object v1, object v2) {
            if (cellValue == null)
                return false;
            switch(@operator)
            {
                case OperatorEnum.NO_COMPARISON:
                    return false;
                case OperatorEnum.BETWEEN:
                    if (cellValue is Double)
                    {
                        // use zero for null
                        double n1 = v1==null? 0:(double)v1;
                        double n2 = v2 == null ? 0 : (double)v2;
                        return (double)cellValue - n1 >= 0 && (double)cellValue - n2 <= 0;
                    }
                    else if (cellValue is String)
                    {
                        String n1 = v1 == null ? "" : (String)v1;
                        String n2 = v2 == null ? "" : (String)v2;
                        return string.Compare((String)cellValue, n1, true) >= 0 && string.Compare((String)cellValue, n2, true) <= 0;
                    }
                    else if (cellValue is Boolean)
                        return false;
                    return false;
                case OperatorEnum.EQUAL:
                    if (v1 == null)
                        return false;
                    if (cellValue is Double)
                    {
                        return (double)cellValue >= 0;
                    }
                    else if (cellValue is String)
                    {
                        return string.Compare((String)cellValue, (String)v1, true) == 0;
                    }
                    else if (cellValue is Boolean)
                    {
                        bool n1 = (bool)v1;
                        return (bool)cellValue==n1;
                    }
                    return false; 
                case OperatorEnum.NOT_EQUAL:
                    if (v1 == null)
                        return true;
                    if (cellValue is String)
                    {
                        String n1 = (String)v1;
                        return string.Compare((String)cellValue, n1, true) != 0;
                    }
                    else if (cellValue is Boolean)
                    {
                        bool n1 = (bool)v1;
                        return (bool)cellValue != n1;
                    }
                    return false;
                case OperatorEnum.GREATER_THAN:
                    if (cellValue is Double)
                    {
                        // use zero for null
                        double n1 = v1 == null ? 0 : (double)v1;
                        return (double)cellValue> n1;
                    }
                    else if (cellValue is String)
                    {
                        String n1 = v1 == null ? "" : (String)v1;
                        return string.Compare((String)cellValue, n1, true) > 0;
                    }
                    else if (cellValue is Boolean)
                    {
                        if (v1 == null)
                            return true;
                        bool n1 = (bool)v1;
                        return ((bool)cellValue).CompareTo(n1)>0;
                    }
                    return false;
                case OperatorEnum.LESS_THAN:
                    if (cellValue is Double)
                    {
                        // use zero for null
                        double n1 = v1 == null ? 0 : (double)v1;
                        return (double)cellValue<n1;
                    }
                    else if (cellValue is String)
                    {
                        if (v1 == null)
                            return false;
                        return string.Compare((String)cellValue, (String)v1, true)< 0;
                    }
                    else if (cellValue is Boolean)
                    {
                        if (v1 == null)
                            return false;
                        bool n1 = (bool)v1;
                        return ((bool)cellValue).CompareTo(n1) < 0;
                    }
                    return false;
                case OperatorEnum.GREATER_OR_EQUAL:
                    if (cellValue is Double)
                    {
                        // use zero for null
                        double n1 = v1 == null ? 0 : (double)v1;
                        return (double)cellValue >= n1;
                    }
                    else if (cellValue is String)
                    {
                        if (v1 == null)
                            return true;
                        return string.Compare((String)cellValue, (String)v1, true) >= 0;
                    }
                    else if (cellValue is Boolean)
                    {
                        if (v1 == null)
                            return false;
                        bool n1 = (bool)v1;
                        return ((bool)cellValue).CompareTo(n1) >= 0;
                    }
                    return false;
                case OperatorEnum.LESS_OR_EQUAL:
                    if (cellValue is Double)
                    {
                        // use zero for null
                        double n1 = v1 == null ? 0 : (double)v1;
                        return (double)cellValue <= n1;
                    }
                    else if (cellValue is String)
                    {
                        if (v1 == null)
                            return false;
                        return string.Compare((String)cellValue, (String)v1, true) <= 0;
                    }
                    else if (cellValue is Boolean)
                    {
                        if (v1 == null)
                            return false;
                        bool n1 = (bool)v1;
                        return ((bool)cellValue).CompareTo(n1) <= 0;
                    }
                    return false;
            }
            return false;

        }
        public static bool IsValidForIncompatibleTypes(OperatorEnum @operator)
        {
            if (@operator == OperatorEnum.NOT_BETWEEN || @operator == OperatorEnum.NOT_EQUAL)
                return true;
            else
                return false;
        }
    }
    public class EvaluationConditionalFormatRule: IComparable<EvaluationConditionalFormatRule>
    {
        /**
    * Note: this class has a natural ordering that is inconsistent with equals.
    */
        protected class ValueAndFormat : IComparable<ValueAndFormat>
        {
            private double? value;
            private String str;
            private String format;
            private DecimalFormat decimalTextFormat;

            public ValueAndFormat(Double value, String format, DecimalFormat df)
            {
                this.value = value;
                this.format = format;
                str = null;
                decimalTextFormat = df;
            }

            public ValueAndFormat(String value, String format)
            {
                this.value = null;
                this.format = format;
                str = value;
                decimalTextFormat = null;
            }

            public bool IsNumber
            {
                get
                {
                    return value != null;
                }
            }

            public double? Value
            {
                get
                {
                    return value;
                }
            }

            public String String
            {
                get
                {
                    return str;
                }
            }

            public override String ToString()
            {
                if (IsNumber)
                {
                    return decimalTextFormat.Format(Value.Value);
                }
                else
                {
                    return this.String;
                }
            }


            public override bool Equals(Object obj)
            {
                if (!(obj is ValueAndFormat))
                {
                    return false;
                }
                ValueAndFormat o = (ValueAndFormat)obj;
                return (value == o.value || value.Equals(o.value))
                        && (format == o.format || format.Equals(o.format))
                        && (str == o.str || str.Equals(o.str));
            }

            /**
             * Note: this class has a natural ordering that is inconsistent with equals.
             * @param o
             * @return value comparison
             */
            public int CompareTo(ValueAndFormat o)
            {
                if (value == null && o.value != null)
                {
                    return 1;
                }
                if (o.value == null && value != null)
                {
                    return -1;
                }
                int cmp = value == null ? 0 : value.Value.CompareTo(o.value);
                if (cmp != 0)
                {
                    return cmp;
                }

                if (str == null && o.str != null)
                {
                    return 1;
                }
                if (o.str == null && str != null)
                {
                    return -1;
                }

                return str == null ? 0 : str.CompareTo(o.str);
            }


            public override int GetHashCode()
            {
                return (str == null ? 0 : str.GetHashCode()) * 37 * 37 + 37 * (value == null ? 0 : value.GetHashCode()) + (format == null ? 0 : format.GetHashCode());
            }

            public static bool operator <(ValueAndFormat left, ValueAndFormat right)
            {
                return left.CompareTo(right) < 0;
            }

            public static bool operator <=(ValueAndFormat left, ValueAndFormat right)
            {
                return left.CompareTo(right) <= 0;
            }

            public static bool operator >(ValueAndFormat left, ValueAndFormat right)
            {
                return left.CompareTo(right) > 0;
            }

            public static bool operator >=(ValueAndFormat left, ValueAndFormat right)
            {
                return left.CompareTo(right) >= 0;
            }
        }

        private WorkbookEvaluator workbookEvaluator;
        private ISheet sheet;
        private IConditionalFormatting formatting;
        private IConditionalFormattingRule rule;

        /* cached values */
        private CellRangeAddress[] regions;
        /**
         * Depending on the rule type, it may want to know about certain values in the region when evaluating {@link #matches(CellReference)},
         * such as top 10, unique, duplicate, average, etc.  This collection stores those if needed so they are not repeatedly calculated
         */
        private Dictionary<CellRangeAddress, List<ValueAndFormat>> meaningfulRegionValues = new Dictionary<CellRangeAddress, List<ValueAndFormat>>();

        private int priority;
        private int formattingIndex;
        private int ruleIndex;
        private String formula1;
        private String formula2;
        private String text;
        // cached for performance, used with cell text comparisons, which are case insensitive and need to be Locale aware (contains, starts with, etc.) 
        private String lowerText;

        private OperatorEnum @operator;
        private ConditionType type;
        // cached for performance, to avoid reading the XMLBean every time a conditionally formatted cell is rendered
        private ExcelNumberFormat numberFormat;
        // cached for performance, used to format numeric cells for string comparisons.  See Bug #61764 for explanation
        private DecimalFormat decimalTextFormat;
        public EvaluationConditionalFormatRule(WorkbookEvaluator workbookEvaluator, ISheet sheet, IConditionalFormatting formatting, int formattingIndex, IConditionalFormattingRule rule, int ruleIndex, CellRangeAddress[] regions)
        {
            this.workbookEvaluator = workbookEvaluator;
            this.sheet = sheet;
            this.formatting = formatting;
            this.rule = rule;
            this.formattingIndex = formattingIndex;
            this.ruleIndex = ruleIndex;

            this.priority = rule.Priority;

            this.regions = regions;
            formula1 = rule.Formula1;
            formula2 = rule.Formula2;

            text = rule.Text;
            lowerText = text == null ? null : text.ToLowerInvariant();

            numberFormat = rule.NumberFormat;
        
            @operator = (OperatorEnum)rule.ComparisonOperation;
            type = rule.ConditionType;
        }
        public ISheet Sheet { get { return sheet; } }
        public IConditionalFormatting Formatting { get { return formatting; } }
        public int FormattingIndex { get { return formattingIndex; } }
        public IConditionalFormattingRule Rule { get { return rule; } } 
        public int RuleIndex { get { return ruleIndex; } }
        public CellRangeAddress[] Regions { get { return regions; } }
        public int Priority { get { return priority; } }
        public string Formula1 { get { return formula1; } }
        public string Formula2 { get { return formula2; } }
        public string Text { get { return text; } }
        public ConditionType Type { get { return type; } }
        public OperatorEnum Operator { get { return @operator; } }
        public ExcelNumberFormat NumberFormat
        {
            get
            {
                return numberFormat;
            }
        }
        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            if (obj.GetType()!=this.GetType())
            {
                return false;
            }
            EvaluationConditionalFormatRule r = (EvaluationConditionalFormatRule)obj;
            return Sheet.SheetName.Equals(r.Sheet.SheetName, StringComparison.InvariantCultureIgnoreCase)
                && FormattingIndex == r.FormattingIndex
                && RuleIndex == r.RuleIndex;
        }
        public int CompareTo(EvaluationConditionalFormatRule o)
        {
            int cmp = Sheet.SheetName.CompareTo(o.Sheet.SheetName);
            if (cmp != 0)
            {
                return cmp;
            }

            int x = Priority;
            int y = o.Priority;
            // logic from Integer.compare()
            cmp = x-y;
            if (cmp != 0)
            {
                return cmp;
            }

            cmp = FormattingIndex - o.FormattingIndex;
            if (cmp != 0)
            {
                return cmp;
            }
            return RuleIndex-o.RuleIndex;
        }
        private ValueEval UnwrapEval(ValueEval eval)
        {
            ValueEval comp = eval;

            while (comp is RefEval) {
                RefEval reference = (RefEval)comp;
                comp = reference.GetInnerValueEval(reference.FirstSheetIndex);
            }
            return comp;
        }
        private bool CheckValue(ICell cell, CellRangeAddress region)
        {
            if (cell == null || DataValidationEvaluator.IsType(cell, CellType.Blank)
               || DataValidationEvaluator.IsType(cell, CellType.Error)
               || (DataValidationEvaluator.IsType(cell, CellType.String)
                       && string.IsNullOrEmpty(cell.StringCellValue)
                   )
               )
            {
                return false;
            }

            ValueEval eval = UnwrapEval(workbookEvaluator.Evaluate(rule.Formula1, ConditionalFormattingEvaluator.GetRef(cell), region));

            String f2 = rule.Formula2;
            ValueEval eval2 = BlankEval.instance;
            if (f2 != null && f2.Length > 0)
            {
                eval2 = UnwrapEval(workbookEvaluator.Evaluate(f2, ConditionalFormattingEvaluator.GetRef(cell), region));
            }
            // we assume the cell has been evaluated, and the current formula value stored
            if (DataValidationEvaluator.IsType(cell, CellType.Boolean)
                    && (eval == BlankEval.instance || eval is BoolEval) 
                && (eval2 == BlankEval.instance || eval2 is BoolEval) 
           ) {
                return OperatorEnumHelper.IsValid(@operator, cell.BooleanCellValue, eval == BlankEval.instance ? (bool?)null : ((BoolEval)eval).BooleanValue, eval2 == BlankEval.instance ? (bool?)null : ((BoolEval)eval2).BooleanValue);
            }
            if (DataValidationEvaluator.IsType(cell, CellType.Numeric)
                    && (eval == BlankEval.instance || eval is NumberEval )
                && (eval2 == BlankEval.instance || eval2 is NumberEval) 
           ) {
                return OperatorEnumHelper.IsValid(@operator, cell.NumericCellValue, eval==BlankEval.instance ? (double?)null : ((NumberEval)eval).NumberValue, eval2 == BlankEval.instance ? (double?)null : ((NumberEval)eval2).NumberValue);
            }
            if (DataValidationEvaluator.IsType(cell, CellType.String)
                    && (eval == BlankEval.instance || eval is StringEval )
                && (eval2 == BlankEval.instance || eval2 is StringEval) 
           ) {
                return OperatorEnumHelper.IsValid(@operator, cell.StringCellValue, eval == BlankEval.instance ? null : ((StringEval)eval).StringValue, eval2 == BlankEval.instance ? null : ((StringEval)eval2).StringValue);
            }

            return OperatorEnumHelper.IsValidForIncompatibleTypes(@operator);
        }
        private bool CheckFormula(CellReference reference, CellRangeAddress region)
        {
            ValueEval comp = UnwrapEval(workbookEvaluator.Evaluate(rule.Formula1, reference, region));

            // Copied for now from DataValidationEvaluator.ValidationEnum.FORMULA#isValidValue()
            if (comp is BlankEval) {
                return true;
            }
            if (comp is ErrorEval) {
                return false;
            }
            if (comp is BoolEval) {
                return ((BoolEval)comp).BooleanValue;
            }
            // empirically tested in Excel - 0=false, any other number = true/valid
            // see test file DataValidationEvaluations.xlsx
            if (comp is NumberEval) {
                return ((NumberEval)comp).NumberValue != 0;
            }
            return false; // anything else is false, such as text
        }
        internal bool Matches(CellReference reference)
        {
            // first check that it is in one of the regions defined for this format
            CellRangeAddress region = null;
            foreach (CellRangeAddress r in regions)
            {
                if (r.IsInRange(reference))
                {
                    region = r;
                    break;
                }
            }

            if (region == null)
            {
                // cell not in range of this rule
                return false;
            }

            ConditionType ruleType = this.Rule.ConditionType;

            // these rules apply to all cells in a region. Specific condition criteria
            // may specify no special formatting for that value partition, but that's display logic
            if (ruleType.Equals(ConditionType.ColorScale)
                || ruleType.Equals(ConditionType.DataBar)
                || ruleType.Equals(ConditionType.IconSet))
            {
                return true;
            }

            ICell cell = null;
            IRow row = sheet.GetRow(reference.Row);
            if (row != null)
            {
                cell = row.GetCell(reference.Col);
            }

            if (ruleType.Equals(ConditionType.CellValueIs))
            {
                // undefined cells never match a VALUE_IS condition
                if (cell == null) return false;
                return CheckValue(cell, region);
            }
            if (ruleType.Equals(ConditionType.Formula))
            {
                return CheckFormula(reference, region);
            }
            if (ruleType.Equals(ConditionType.Filter))
            {
                return CheckFilter(cell, reference, region);
            }
            // TODO: anything else, we don't handle yet, such as top 10
            return false;
        }
        private bool CheckFilter(ICell cell, CellReference reference, CellRangeAddress region)
        {
            ConditionFilterType? filterType = rule.ConditionFilterType;
            if (filterType == null)
            {
                return false;
            }
            ValueAndFormat cv = GetCellValue(cell);
            switch (filterType)
            {
                case ConditionFilterType.FILTER:
                    return false; // we don't evaluate HSSF filters yet
                case ConditionFilterType.TOP_10:
                    // from testing, Excel only operates on numbers and dates (which are stored as numbers) in the range.
                    // numbers stored as text are ignored, but numbers formatted as text are treated as numbers.

                    return GetMeaningfulValues(region, true,  (allValues)=> {
                        IConditionFilterData fc = rule.FilterConfiguration;

                        if (!fc.Bottom)
                        {
                            allValues.Sort();
                            allValues.Reverse();
                        }
                        else
                        {
                            allValues.Sort();
                        }

                        int limit = (int)fc.Rank;
                        if (fc.Percent)
                        {
                            limit = allValues.Count * limit / 100;
                        }
                        if (allValues.Count <= limit)
                        {
                            return allValues;
                        }
                        return allValues.GetRange(0, limit);
                    }).Contains(cv);
                case ConditionFilterType.UNIQUE_VALUES:
                // Per Excel help, "duplicate" means matching value AND format
                // https://support.office.com/en-us/article/Filter-for-unique-values-or-remove-duplicate-values-ccf664b0-81d6-449b-bbe1-8daaec1e83c2
                 return GetMeaningfulValues(region, true, (allValues) => {
                     allValues.Sort();
                     List<ValueAndFormat> unique = new List<ValueAndFormat>();

                     for (int i = 0; i < allValues.Count; i++)
                     {
                         ValueAndFormat v = allValues[i];
                         // skip this if the current value matches the next one, or is the last one and matches the previous one
                         if ((i < allValues.Count - 1 && v.Equals(allValues[i + 1])) || (i > 0 && i == allValues.Count - 1 && v.Equals(allValues[i - 1])))
                         {
                             // current value matches next value, skip both
                             i++;
                             continue;
                         }
                         unique.Add(v);
                     }

                     return unique;
                 }).Contains(cv);
                case ConditionFilterType.DUPLICATE_VALUES:
                    // Per Excel help, "duplicate" means matching value AND format
                    // https://support.office.com/en-us/article/Filter-for-unique-values-or-remove-duplicate-values-ccf664b0-81d6-449b-bbe1-8daaec1e83c2
                    return GetMeaningfulValues(region, true, (allValues) => {
                        allValues.Sort();
                        List<ValueAndFormat> dup = new List<ValueAndFormat>();

                        for (int i = 0; i < allValues.Count; i++)
                        {
                            ValueAndFormat v = allValues[i];
                            // skip this if the current value matches the next one, or is the last one and matches the previous one
                            if ((i < allValues.Count - 1 && v.Equals(allValues[i + 1])) || (i > 0 && i == allValues.Count - 1 && v.Equals(allValues[i - 1])))
                            {
                                // current value matches next value, add one
                                dup.Add(v);
                                i++;
                            }
                        }
                        return dup;
                    }).Contains(cv);
                case ConditionFilterType.ABOVE_AVERAGE:
                    // from testing, Excel only operates on numbers and dates (which are stored as numbers) in the range.
                    // numbers stored as text are ignored, but numbers formatted as text are treated as numbers.

                    IConditionFilterData conf = rule.FilterConfiguration;

                    // actually ordered, so iteration order is predictable
                    List<ValueAndFormat> values = GetMeaningfulValues(region, false, (allValues) => {
                        double total = 0;
                        ValueEval[] pop = new ValueEval[allValues.Count];
                        for (int i = 0; i < allValues.Count; i++)
                        {
                            ValueAndFormat v = allValues[i];
                            total += (double)v.Value;
                            pop[i] = new NumberEval((double)v.Value);
                        }

                        List<ValueAndFormat> avgSet = new List<ValueAndFormat>(1);
                        avgSet.Add(new ValueAndFormat(allValues.Count == 0 ? 0 : total / allValues.Count, null, decimalTextFormat));

                        double stdDev2 = allValues.Count <= 1 ? 0 : ((NumberEval)AggregateFunction.STDEV.Evaluate(pop, 0, 0)).NumberValue;
                        avgSet.Add(new ValueAndFormat(stdDev2, null, decimalTextFormat));
                        return avgSet;
                    });
                    double? val = cv.IsNumber ? cv.Value : null;
                    if (val == null)
                    {
                        return false;
                    }

                    double avg = (double)values[0].Value;
                    double stdDev = (double)values[1].Value;

                    /*
                     * use StdDev, aboveAverage, equalAverage to find:
                     * comparison value
                     * operator type
                     */

                    double comp = conf.StdDev > 0 ? (avg + (conf.AboveAverage ? 1 : -1) * stdDev * conf.StdDev) : avg;

                    OperatorEnum op;
                    if (conf.AboveAverage)
                    {
                        if (conf.EqualAverage)
                        {
                            op = OperatorEnum.GREATER_OR_EQUAL;
                        }
                        else
                        {
                            op = OperatorEnum.GREATER_THAN;
                        }
                    }
                    else
                    {
                        if (conf.EqualAverage)
                        {
                            op = OperatorEnum.LESS_OR_EQUAL;
                        }
                        else
                        {
                            op = OperatorEnum.LESS_THAN;
                        }
                    }
                    return OperatorEnumHelper.IsValid(op, val, comp, null);
                case ConditionFilterType.CONTAINS_TEXT:
                    // implemented both by a cfRule "text" attribute and a formula.  Use the text.
                    return text == null ? false : cv.ToString().ToLowerInvariant().Contains(lowerText);
                case ConditionFilterType.NOT_CONTAINS_TEXT:
                    // implemented both by a cfRule "text" attribute and a formula.  Use the text.
                    return text == null ? true : !cv.ToString().ToLowerInvariant().Contains(lowerText);
                case ConditionFilterType.BEGINS_WITH:
                    // implemented both by a cfRule "text" attribute and a formula.  Use the text.
                    return cv.ToString().ToLowerInvariant().StartsWith(lowerText);
                case ConditionFilterType.ENDS_WITH:
                    // implemented both by a cfRule "text" attribute and a formula.  Use the text.
                    return cv.ToString().ToLowerInvariant().EndsWith(lowerText);
                case ConditionFilterType.CONTAINS_BLANKS:
                    try
                    {
                        String v = cv.String;
                        // see TextFunction.TRIM for implementation
                        return v == null || v.Trim().Length == 0;
                    }
                    catch
                    {
                        // not a valid string value, and not a blank cell (that's checked earlier)
                        return false;
                    }
                case ConditionFilterType.NOT_CONTAINS_BLANKS:
                    try
                    {
                        String v = cv.String;
                        // see TextFunction.TRIM for implementation
                        return v != null && v.Trim().Length > 0;
                    }
                    catch
                    {
                        // not a valid string value, but not blank
                        return true;
                    }
                case ConditionFilterType.CONTAINS_ERRORS:
                    return cell != null && DataValidationEvaluator.IsType(cell, CellType.Error);
                case ConditionFilterType.NOT_CONTAINS_ERRORS:
                    return cell == null || !DataValidationEvaluator.IsType(cell, CellType.Error);
                case ConditionFilterType.TIME_PERIOD:
                    // implemented both by a cfRule "text" attribute and a formula.  Use the formula.
                    return CheckFormula(reference, region);
                default:
                    return false;
            }
        }
        private List<ValueAndFormat> GetMeaningfulValues(CellRangeAddress region, bool withText, Func<List<ValueAndFormat>, List<ValueAndFormat>> func)
        {
            if (meaningfulRegionValues.ContainsKey(region))
            {
                return meaningfulRegionValues[region];
            }
            List<ValueAndFormat> values = new List<ValueAndFormat>();
            List<ValueAndFormat> allValues = new List<ValueAndFormat>((region.LastColumn - region.FirstColumn + 1) * (region.LastRow - region.FirstRow + 1));

            for (int r = region.FirstRow; r <= region.LastRow; r++)
            {
                IRow row = sheet.GetRow(r);
                if (row == null)
                {
                    continue;
                }
                for (int c = region.FirstColumn; c <= region.LastColumn; c++)
                {
                    ICell cell = row.GetCell(c);
                    ValueAndFormat cv = GetCellValue(cell);
                    if (withText || cv.IsNumber)
                    {
                        allValues.Add(cv);
                    }
                }
            }
            values = func(allValues);
            meaningfulRegionValues.Add(region, values);

            return values;
        }
        private ValueAndFormat GetCellValue(ICell cell)
        {
            if (cell != null)
            {
                String format = cell.CellStyle.GetDataFormatString();
                CellType type = cell.CellType;
                if (type == CellType.Formula)
                {
                    type = cell.CachedFormulaResultType;
                }
                switch (type)
                {
                    case CellType.Numeric:
                        return new ValueAndFormat(cell.NumericCellValue, format, decimalTextFormat);
                    case CellType.String:
                    case CellType.Boolean:
                        return new ValueAndFormat(cell.StringCellValue, format);
                    default:
                        break;
                }
            }
            return new ValueAndFormat("", "");
        }
        public override int GetHashCode()
        {
            int hash = sheet.SheetName.GetHashCode();
            hash = 31 * hash + formattingIndex;
            hash = 31 * hash + ruleIndex;
            return hash;
        }
    }
}
