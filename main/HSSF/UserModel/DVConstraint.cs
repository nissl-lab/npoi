/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

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
    using NPOI.HSSF.Model;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula;
    using NPOI.SS.UserModel;
    using System.Text;
    using NPOI.SS.Util;
    using System.Globalization;
using NPOI.HSSF.Record;

    /**
     * 
     * @author Josh Micich
     */
    public class DVConstraint : IDataValidationConstraint
    {
        /* package */
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

        // convenient access to ValidationType namespace
        //private static ValidationType VT = null;


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
            : this(ValidationType.LIST, OperatorType.IGNORED,
                listFormula, null, Double.NaN, Double.NaN, excplicitListValues)
        {
            ;
        }

        /**
         * Creates a number based data validation constraint. The text values entered for expr1 and expr2
         * can be either standard Excel formulas or formatted number values. If the expression starts 
         * with '=' it is Parsed as a formula, otherwise it is Parsed as a formatted number. 
         * 
         * @param validationType one of {@link NPOI.SS.UserModel.DataValidationConstraint.ValidationType#ANY},
         * {@link NPOI.SS.UserModel.DataValidationConstraint.ValidationType#DECIMAL},
         * {@link NPOI.SS.UserModel.DataValidationConstraint.ValidationType#INTEGER},
         * {@link NPOI.SS.UserModel.DataValidationConstraint.ValidationType#TEXT_LENGTH}
         * @param comparisonOperator any constant from {@link NPOI.SS.UserModel.DataValidationConstraint.OperatorType} enum
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
            String formula1 = GetFormulaFromTextExpression(expr1);
            Double value1 = formula1 == null ? ConvertNumber(expr1) : double.NaN;
            // formula2 and value2 are mutually exclusive
            String formula2 = GetFormulaFromTextExpression(expr2);
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
         * with '=' it is Parsed as a formula, otherwise it is Parsed as a formatted time.  To parse 
         * formatted times, two formats are supported:  "HH:MM" or "HH:MM:SS".  This is contrary to 
         * Excel which uses the default time format from the OS.
         * 
         * @param comparisonOperator constant from {@link NPOI.SS.UserModel.DataValidationConstraint.OperatorType} enum
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
            String formula1 = GetFormulaFromTextExpression(expr1);
            Double value1 = formula1 == null ? ConvertTime(expr1) : Double.NaN;
            // formula2 and value2 are mutually exclusive
            String formula2 = GetFormulaFromTextExpression(expr2);
            Double value2 = formula2 == null ? ConvertTime(expr2) : Double.NaN;
            return new DVConstraint(ValidationType.TIME, comparisonOperator, formula1, formula2, value1, value2, null);
        }

        /**
         * Creates a date based data validation constraint. The text values entered for expr1 and expr2
         * can be either standard Excel formulas or formatted date values. If the expression starts 
         * with '=' it is Parsed as a formula, otherwise it is Parsed as a formatted date (Excel uses 
         * the same convention).  To parse formatted dates, a date format needs to be specified.  This
         * is contrary to Excel which uses the default short date format from the OS.
         * 
         * @param comparisonOperator constant from {@link NPOI.SS.UserModel.DataValidationConstraint.OperatorType} enum
         * @param expr1 date formula (when first char is '=') or formatted date value
         * @param expr2 date formula (when first char is '=') or formatted date value
         * @param dateFormat ignored if both expr1 and expr2 are formulas.  Default value is "YYYY/MM/DD"
         * otherwise any other valid argument for <c>SimpleDateFormat</c> can be used
         * @see <a href='http://java.sun.com/j2se/1.5.0/docs/api/java/text/DateFormat.html'>SimpleDateFormat</a>
         */
        public static DVConstraint CreateDateConstraint(int comparisonOperator, String expr1, String expr2, String dateFormat)
        {
            if (expr1 == null)
            {
                throw new ArgumentException("expr1 must be supplied");
            }
            OperatorType.ValidateSecondArg(comparisonOperator, expr2);
            SimpleDateFormat df = dateFormat == null ? null : new SimpleDateFormat(dateFormat);

            // formula1 and value1 are mutually exclusive
            String formula1 = GetFormulaFromTextExpression(expr1);
            Double value1 = formula1 == null ? ConvertDate(expr1, df) : Double.NaN;
            // formula2 and value2 are mutually exclusive
            String formula2 = GetFormulaFromTextExpression(expr2);
            Double value2 = formula2 == null ? ConvertDate(expr2, df) : Double.NaN;
            return new DVConstraint(ValidationType.DATE, comparisonOperator, formula1, formula2, value1, value2, null);
        }

        /**
         * Distinguishes formula expressions from simple value expressions.  This logic is only 
         * required by a few factory methods in this class that create data validation constraints
         * from more or less the same parameters that would have been entered in the Excel UI.  The
         * data validation dialog box uses the convention that formulas begin with '='.  Other methods
         * in this class follow the POI convention (formulas and values are distinct), so the '=' 
         * convention is not used there.
         *  
         * @param textExpr a formula or value expression
         * @return all text After '=' if textExpr begins with '='. Otherwise <code>null</code> if textExpr does not begin with '='
         */
        private static String GetFormulaFromTextExpression(String textExpr)
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
         * @return <code>null</code> if numberStr is <code>null</code>
         */
        private static Double ConvertNumber(String numberStr)
        {
            if (numberStr == null)
            {
                return Double.NaN;
            }
            try
            {
                return double.Parse(numberStr, CultureInfo.CurrentCulture);
            }
            catch (FormatException)
            {
                throw new InvalidOperationException("The supplied text '" + numberStr
                        + "' could not be parsed as a number");
            }
        }

        /**
         * @return <code>null</code> if timeStr is <code>null</code>
         */
        private static Double ConvertTime(String timeStr)
        {
            if (timeStr == null)
            {
                return Double.NaN;
            }
            return HSSFDateUtil.ConvertTime(timeStr);
        }
        /**
         * @param dateFormat pass <code>null</code> for default YYYYMMDD
         * @return <code>null</code> if timeStr is <code>null</code>
         */
        private static Double ConvertDate(String dateStr, SimpleDateFormat dateFormat)
        {
            if (dateStr == null)
            {
                return Double.NaN;
            }
            DateTime dateVal;
            if (dateFormat == null)
            {
                dateVal = HSSFDateUtil.ParseYYYYMMDDDate(dateStr);
            }
            else
            {
                try
                {
                    dateVal = DateTime.Parse(dateStr, CultureInfo.CurrentCulture);
                }
                catch (FormatException e)
                {
                    throw new InvalidOperationException("Failed to parse date '" + dateStr
                            + "' using specified format '" + dateFormat + "'", e);
                }
            }
            return HSSFDateUtil.GetExcelDate(dateVal);
        }

        public static DVConstraint CreateCustomFormulaConstraint(String formula)
        {
            if (formula == null)
            {
                throw new ArgumentException("formula must be supplied");
            }
            return new DVConstraint(ValidationType.FORMULA, OperatorType.IGNORED, formula, null, double.NaN, double.NaN, null);
        }

        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidationConstraint#getValidationType()
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

        public int Operator
        {
            get
            {
                return _operator;
            }
            set
            {
                _operator = value;
            }
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
        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidationConstraint#getFormula1()
         */
        public String Formula1
        {
            get
            {
                return _formula1;
            }
            set
            {
                _value1 = double.NaN;
                _explicitListValues = null;
                _formula1 = value;
            }
        }

        /* (non-Javadoc)
         * @see NPOI.HSSF.UserModel.DataValidationConstraint#getFormula2()
         */
        public String Formula2
        {
            get
            {
                return _formula2;
            }
            set
            {
                _value2 = double.NaN;
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

        /**
         * @return both Parsed formulas (for expression 1 and 2). 
         */
        /* package */
        public FormulaPair CreateFormulas(HSSFSheet sheet)
        {
            Ptg[] formula1;
            Ptg[] formula2;
            if (IsListValidationType)
            {
                formula1 = CreateListFormula(sheet);
                formula2 = Ptg.EMPTY_PTG_ARRAY;
            }
            else
            {
                formula1 = ConvertDoubleFormula(_formula1, _value1, sheet);
                formula2 = ConvertDoubleFormula(_formula2, _value2, sheet);
            }
            return new FormulaPair(formula1, formula2);
        }

        private Ptg[] CreateListFormula(HSSFSheet sheet)
        {

            if (_explicitListValues == null)
            {
                IWorkbook wb = sheet.Workbook;
                // formula is Parsed with slightly different RVA rules: (root node type must be 'reference')
                return HSSFFormulaParser.Parse(_formula1, (HSSFWorkbook)wb, FormulaType.DataValidationList, wb.GetSheetIndex(sheet));
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
         * @return The Parsed token array representing the formula or value specified. 
         * Empty array if both formula and value are <code>null</code>
         */
        private static Ptg[] ConvertDoubleFormula(String formula, Double value, HSSFSheet sheet)
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
            IWorkbook wb = sheet.Workbook;
            return HSSFFormulaParser.Parse(formula, (HSSFWorkbook)wb, FormulaType.Cell, wb.GetSheetIndex(sheet));
        }

        internal static DVConstraint CreateDVConstraint(DVRecord dvRecord, IFormulaRenderingWorkbook book)
        {
            switch (dvRecord.DataType)
            {
                case ValidationType.ANY:
                    return new DVConstraint(ValidationType.ANY, dvRecord.ConditionOperator, null, null, double.NaN, double.NaN, null);
                case ValidationType.INTEGER:
                case ValidationType.DECIMAL:
                case ValidationType.DATE:
                case ValidationType.TIME:
                case ValidationType.TEXT_LENGTH:
                    FormulaValuePair pair1 = toFormulaString(dvRecord.Formula1, book);
                    FormulaValuePair pair2 = toFormulaString(dvRecord.Formula2, book);
                    return new DVConstraint(dvRecord.DataType, dvRecord.ConditionOperator, pair1.formula(),
                            pair2.formula(), pair1.Value, pair2.Value, null);
                case ValidationType.LIST:
                    if (dvRecord.ListExplicitFormula)
                    {
                        String values = toFormulaString(dvRecord.Formula1, book).AsString();
                        if (values.StartsWith("\""))
                        {
                            values = values.Substring(1);
                        }
                        if (values.EndsWith("\""))
                        {
                            values = values.Substring(0, values.Length - 1);
                        }
                        String[] explicitListValues = values.Split("\0".ToCharArray());
                        return CreateExplicitListConstraint(explicitListValues);
                    }
                    else
                    {
                        String listFormula = toFormulaString(dvRecord.Formula1, book).AsString();
                        return CreateFormulaListConstraint(listFormula);
                    }
                case ValidationType.FORMULA:
                    return CreateCustomFormulaConstraint(toFormulaString(dvRecord.Formula1, book).AsString());
                default:
                    throw new InvalidOperationException(string.Format("validationType={0}", dvRecord.DataType));
            }
        }

        private class FormulaValuePair
        {
            internal String _formula;
            internal String _value;

            public String formula()
            {
                return _formula;
            }

            public Double Value
            {
                get
                {
                    if (_value == null)
                    {
                        return double.NaN;
                    }
                    return double.Parse(_value);
                }
            }

            public String AsString()
            {
                if (_formula != null)
                {
                    return _formula;
                }
                if (_value != null)
                {
                    return _value;
                }
                return null;
            }
        }

        private static FormulaValuePair toFormulaString(Ptg[] ptgs, IFormulaRenderingWorkbook book)
        {
            FormulaValuePair pair = new FormulaValuePair();
            if (ptgs != null && ptgs.Length > 0)
            {
                String aString = FormulaRenderer.ToFormulaString(book, ptgs);
                if (ptgs.Length == 1 && ptgs[0].GetType() == typeof(NumberPtg))
                {
                    pair._value = aString;
                }
                else
                {
                    pair._formula = aString;
                }
            }
            return pair;
        }
    }

}