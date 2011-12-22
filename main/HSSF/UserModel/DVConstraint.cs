/* ====================================================================
   Copyright 2002-2004   Apache Software Foundation

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Text;
    using System.Collections;
    using System.Text.RegularExpressions;
    using NPOI.HSSF.Model;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.PTG;
    using NPOI.HSSF.Record;

    public class FormulaPair
    {

        private Ptg[] _formula1;
        private Ptg[] _formula2;

        public FormulaPair(Ptg[] formula1, Ptg[] formula2)
        {
            _formula1 = formula1;
            _formula2 = formula2;
        }
        public Ptg[] Formula1
        {
            get
            {
                return _formula1;
            }
        }
        public Ptg[] Formula2
        {
            get
            {
                return _formula2;
            }
        }

    }
    /**
     * 
     * @author Josh Micich
     */
    public class DVConstraint
    {
        /**
         * ValidationType enum
         */
        public class ValidationType
        {
            /** 'Any value' type - value not restricted */
            public const int ANY = 0x00;
            /** Integer ('Whole number') type */
            public const int INTEGER = 0x01;
            /** Decimal type */
            public const int DECIMAL = 0x02;
            /** List type ( combo box type ) */
            public const int LIST = 0x03;
            /** Date type */
            public const int DATE = 0x04;
            /** Time type */
            public const int TIME = 0x05;
            /** String length type */
            public const int TEXT_LENGTH = 0x06;
            /** Formula ( 'Custom' ) type */
            public const int FORMULA = 0x07;
        }
        /**
         * Condition operator enum
         */
        public class OperatorType
        {
            private OperatorType()
            {
                // no instances of this class
            }

            public const int BETWEEN = 0x00;
            public const int NOT_BETWEEN = 0x01;
            public const int EQUAL = 0x02;
            public const int NOT_EQUAL = 0x03;
            public const int GREATER_THAN = 0x04;
            public const int LESS_THAN = 0x05;
            public const int GREATER_OR_EQUAL = 0x06;
            public const int LESS_OR_EQUAL = 0x07;
            /** default value to supply when the operator type is not used */
            public const int IGNORED = BETWEEN;

            /* package */
            internal static void ValidateSecondArg(int comparisonOperator, String paramValue)
            {
                switch (comparisonOperator)
                {
                    case BETWEEN:
                    case NOT_BETWEEN:
                        if (paramValue == null)
                        {
                            throw new ArgumentException("expr2 must be supplied for 'between' comparisons");
                        }
                        break;
                    // all other operators don't need second arg
                }
            }
        }


        private int _validationType;
        private int _operator;
        private String[] _explicitListValues;

        private String _formula1;
        private String _formula2;
        private Double _value1;
        private Double _value2;


        private DVConstraint(int validationType, int comparisonOperator, String formulaA,
                String formulaB, Double value1, Double value2, String[] excplicitListValues)
        {
            _validationType = validationType;
            _operator = comparisonOperator;
            _formula1 = formulaA;
            _formula2 = formulaB;
            _value1 = value1;
            _value2 = value2;
            _explicitListValues = excplicitListValues;
        }


        /**
         * Creates a list constraint
         */
        private DVConstraint(String listFormula, String[] excplicitListValues)
            :  this(ValidationType.LIST, OperatorType.IGNORED,
                listFormula, null, Double.NaN, Double.NaN, excplicitListValues)
        {

        }

        /**
         * Creates a number based data validation constraint. The text values entered for expr1 and expr2
         * can be either standard Excel formulas or formatted number values. If the expression starts 
         * with '=' it is parsed as a formula, otherwise it is parsed as a formatted number. 
         * 
         * @param validationType one of {@link ValidationType#ANY}, {@link ValidationType#DECIMAL},
         * {@link ValidationType#INTEGER}, {@link ValidationType#TEXT_LENGTH}
         * @param comparisonOperator any constant from {@link OperatorType} enum
         * @param expr1 date formula (when first char is '=') or formatted number value
         * @param expr2 date formula (when first char is '=') or formatted number value
         */
        public static DVConstraint CreateNumericConstraint(int validationType, int comparisonOperator,
                String expr1, String expr2)
        {
            switch (validationType)
            {
                case ValidationType.ANY:
                    if (expr1 != null || expr2 != null)
                    {
                        throw new ArgumentException("expr1 and expr2 must be null for validation type 'any'");
                    }
                    break;
                case ValidationType.DECIMAL:
                case ValidationType.INTEGER:
                case ValidationType.TEXT_LENGTH:
                    if (expr1 == null)
                    {
                        throw new ArgumentException("expr1 must be supplied");
                    }
                    OperatorType.ValidateSecondArg(comparisonOperator, expr2);
                    break;
                default:
                    throw new ArgumentException("Validation Type ("
                            + validationType + ") not supported with this method");
            }
            // formula1 and value1 are mutually exclusive
            String formula1 = getFormulaFromTextExpression(expr1);
            Double value1 = formula1 == null ? ConvertNumber(expr1) : double.NaN;
            // formula2 and value2 are mutually exclusive
            String formula2 = getFormulaFromTextExpression(expr2);
            Double value2 = formula2 == null ? ConvertNumber(expr2) : double.NaN;
            return new DVConstraint(validationType, comparisonOperator, formula1, formula2, value1, value2, null);
        }

        public static DVConstraint CreateFormulaListConstraint(String listFormula)
        {
            return new DVConstraint(listFormula, null);
        }
        public static DVConstraint CreateExplicitListConstraint(String[] explicitListValues)
        {
            return new DVConstraint(null, explicitListValues);
        }


        /**
         * Creates a time based data validation constraint. The text values entered for expr1 and expr2
         * can be either standard Excel formulas or formatted time values. If the expression starts 
         * with '=' it is parsed as a formula, otherwise it is parsed as a formatted time.  To parse 
         * formatted times, two formats are supported:  "HH:MM" or "HH:MM:SS".  This is contrary to 
         * Excel which uses the default time format from the OS.
         * 
         * @param comparisonOperator constant from {@link OperatorType} enum
         * @param expr1 date formula (when first char is '=') or formatted time value
         * @param expr2 date formula (when first char is '=') or formatted time value
         */
        public static DVConstraint CreateTimeConstraint(int comparisonOperator, String expr1, String expr2)
        {
            if (expr1 == null)
            {
                throw new ArgumentException("expr1 must be supplied");
            }
            OperatorType.ValidateSecondArg(comparisonOperator, expr1);

            // formula1 and value1 are mutually exclusive
            String formula1 = getFormulaFromTextExpression(expr1);
            Double value1 = formula1 == null ? ConvertTime(expr1) :double.NaN;
            // formula2 and value2 are mutually exclusive
            String formula2 = getFormulaFromTextExpression(expr2);
            Double value2 = formula2 == null ? ConvertTime(expr2) : double.NaN;
            return new DVConstraint(ValidationType.TIME, comparisonOperator, formula1, formula2, value1, value2, null);

        }
        /**
         * Creates a date based data validation constraint. The text values entered for expr1 and expr2
         * can be either standard Excel formulas or formatted date values. If the expression starts 
         * with '=' it is parsed as a formula, otherwise it is parsed as a formatted date (Excel uses 
         * the same convention).  To parse formatted dates, a date format needs to be specified.  This
         * is contrary to Excel which uses the default short date format from the OS.
         * 
         * @param comparisonOperator constant from {@link OperatorType} enum
         * @param expr1 date formula (when first char is '=') or formatted date value
         * @param expr2 date formula (when first char is '=') or formatted date value
         * @param dateFormat ignored if both expr1 and expr2 are formulas.  Default value is "YYYY/MM/DD"
         * otherwise any other valid argument for <tt>SimpleDateFormat</tt> can be used
         * @see <a href='http://java.sun.com/j2se/1.5.0/docs/api/java/text/DateFormat.html'>SimpleDateFormat</a>
         */
        public static DVConstraint CreateDateConstraint(int comparisonOperator, String expr1, String expr2, String dateFormat)
        {
            if (expr1 == null)
            {
                throw new ArgumentException("expr1 must be supplied");
            }
            OperatorType.ValidateSecondArg(comparisonOperator, expr2);
            

            // formula1 and value1 are mutually exclusive
            String formula1 = getFormulaFromTextExpression(expr1);
            Double value1 = formula1 == null ? ConvertDate(expr1, dateFormat) : Double.NaN;
            // formula2 and value2 are mutually exclusive
            String formula2 = getFormulaFromTextExpression(expr2);
            Double value2 = formula2 == null ? ConvertDate(expr2, dateFormat) : Double.NaN;
            return new DVConstraint(ValidationType.DATE, comparisonOperator, formula1, formula2, value1, value2, null);
        }

        /**
         * Distinguishes formula expressions from simple value expressions.  This logic is only 
         * required by a few factory methods in this class that Create data validation constraints
         * from more or less the same parameters that would have been entered in the Excel UI.  The
         * data validation dialog box uses the convention that formulas begin with '='.  Other methods
         * in this class follow the POI convention (formulas and values are distinct), so the '=' 
         * convention is not used there.
         *  
         * @param textExpr a formula or value expression
         * @return all text after '=' if textExpr begins with '='. Otherwise <c>null</c> if textExpr does not begin with '='
         */
        private static String getFormulaFromTextExpression(String textExpr)
        {
            if (textExpr == null)
            {
                return null;
            }
            if (textExpr.Length < 1)
            {
                throw new ArgumentException("Empty string is not a valid formula/value expression");
            }
            if (textExpr[0] == '=')
            {
                return textExpr.Substring(1);
            }
            return null;
        }


        /**
         * @return <c>null</c> if numberStr is <c>null</c>
         */
        private static Double ConvertNumber(String numberStr)
        {
            if (numberStr == null)
            {
                return Double.NaN;
            }
            try
            {
                return Double.Parse(numberStr);
            }
            catch (FormatException)
            {
                throw new InvalidOperationException("The supplied text '" + numberStr
                        + "' could not be parsed as a number");
            }
        }

        /**
         * @return <c>null</c> if timeStr is <c>null</c>
         */
        private static Double ConvertTime(String timeStr)
        {
            if (timeStr == null)
            {
                return Double.NaN;
            }
            return NPOI.SS.UserModel.DateUtil.ConvertTime(timeStr);
        }
        /**
         * @param dateFormat pass <c>null</c> for default YYYYMMDD
         * @return <c>null</c> if timeStr is <c>null</c>
         */
        private static Double ConvertDate(String dateStr, string dateFormat)
        {
            if (dateStr == null)
            {
                return Double.NaN;
            }
            DateTime dateVal;
            if (dateFormat == null)
            {
                dateVal = DateTime.Parse(dateStr);
            }
            else
            {
                try
                {
                    dateVal = DateTime.Parse(dateStr);
                }
                catch (FormatException e)
                {
                    throw new InvalidOperationException("Failed to parse date '" + dateStr
                            + "' using specified format '" + dateFormat + "'", e);
                }
            }
            return NPOI.SS.UserModel.DateUtil.GetExcelDate(dateVal);
        }

        public static DVConstraint CreateCustomFormulaConstraint(String formula)
        {
            if (formula == null)
            {
                throw new ArgumentException("formula must be supplied");
            }
            return new DVConstraint(ValidationType.FORMULA, OperatorType.IGNORED, formula, null, Double.NaN, Double.NaN, null);
        }

        /**
         * @return both parsed formulas (for expression 1 and 2). 
         */
        /* package */
        public FormulaPair CreateFormulas(HSSFWorkbook workbook)
        {
            Ptg[] formula1;
            Ptg[] formula2;
            if (IsListValidationType)
            {
                formula1 = CreateListFormula(workbook);
                formula2 = Ptg.EMPTY_PTG_ARRAY;
            }
            else
            {
                formula1 = ConvertDoubleFormula(_formula1, _value1, workbook);
                formula2 = ConvertDoubleFormula(_formula2, _value2, workbook);
            }
            return new FormulaPair(formula1, formula2);
        }

        private Ptg[] CreateListFormula(HSSFWorkbook workbook)
        {

            if (_explicitListValues == null)
            {
                // formula is parsed with slightly different RVA rules: (root node type must be 'reference')
                return HSSFFormulaParser.Parse(_formula1, workbook, FormulaType.DATAVALIDATION_LIST);
                // To do: Excel places restrictions on the available operations within a list formula.
                // Some things like union and intersection are not allowed.
            }
            // explicit list was provided
            StringBuilder sb = new StringBuilder(_explicitListValues.Length * 16);
            for (int i = 0; i < _explicitListValues.Length; i++)
            {
                if (i > 0)
                {
                    sb.Append('\0'); // list delimiter is the nul char
                }
                sb.Append(_explicitListValues[i]);

            }
            return new Ptg[] { new StringPtg(sb.ToString()), };
        }

        /**
         * @return The parsed token array representing the formula or value specified. 
         * Empty array if both formula and value are <c>null</c>
         */
        private static Ptg[] ConvertDoubleFormula(String formula, Double value, HSSFWorkbook workbook)
        {
            if (formula == null)
            {
                if (double.IsNaN(value))
                {
                    return Ptg.EMPTY_PTG_ARRAY;
                }
                return new Ptg[] { new NumberPtg(value), };
            }
            if (!double.IsNaN(value))
            {
                throw new InvalidOperationException("Both formula and value cannot be present");
            }
            return HSSFFormulaParser.Parse(formula, workbook);
        }


        /**
         * @return data validation type of this constraint
         * @see ValidationType
         */
        public int GetValidationType()
        {
               return _validationType;
        }
        /**
         * Convenience method
         * @return <c>true</c> if this constraint is a 'list' validation
         */
        public bool IsListValidationType
        {
            get
            {
                return _validationType == ValidationType.LIST;
            }
        }
        /**
         * Convenience method
         * @return <c>true</c> if this constraint is a 'list' validation with explicit values
         */
        public bool IsExplicitList
        {
            get
            {
                return _validationType == ValidationType.LIST && _explicitListValues != null;
            }
        }
        /**
         * @return the operator used for this constraint
         * @see OperatorType
         */
        public int Operator
        {
            get
            {
                return _operator;
            }
            set { _operator = value; }
        }

        public String[] ExplicitListValues
        {
            get
            {
                return _explicitListValues;
            }
            set
            {
                if (_validationType != ValidationType.LIST)
                {
                    throw new InvalidOperationException("Cannot setExplicitListValues on non-list constraint");
                }
                _formula1 = null;
                _explicitListValues = value;
            }
        }
        /**
         * @return the formula for expression 1. May be <c>null</c>
         */
        public String Formula1
        {
            get
            {
                return _formula1;
            }
            set
            {
                _value1 = Double.NaN;
                _explicitListValues = null;
                _formula1 = value;
            }
        }
        /**
         * @return the formula for expression 2. May be <c>null</c>
         */
        public String Formula2
        {
            get
            {
                return _formula2;
            }
            set
            {
                _value2 = Double.NaN;
                _formula2 = value;
            }
        }

        /**
         * @return the numeric value for expression 1. May be <c>null</c>
         */
        public Double Value1
        {
            get
            {
                return _value1;
            }
            set
            {
                _formula1 = null;
                _value1 = value;
            }
        }


        /**
         * @return the numeric value for expression 2. May be <c>null</c>
         */
        public Double Value2
        {
            get
            {
                return _value2;
            }
            set
            {
                _formula2 = null;
                _value2 = value;
            }
        }

    }
}
