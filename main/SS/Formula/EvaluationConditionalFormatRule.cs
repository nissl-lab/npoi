using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.SS.Formula
{
    public enum OperatorEnum
    {
        NO_COMPARISON,
        BETWEEN,
        NOT_BETWEEN,
        EQUAL,
        NOT_EQUAL,
        GREATER_THAN,
        LESS_THAN,
        GREATER_OR_EQUAL,
        LESS_OR_EQUAL
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
        public override int GetHashCode()
        {
            int hash = sheet.SheetName.GetHashCode();
            hash = 31 * hash + formattingIndex;
            hash = 31 * hash + ruleIndex;
            return hash;
        }
    }
}
