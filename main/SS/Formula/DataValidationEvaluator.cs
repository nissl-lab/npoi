/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NPOI.SS.Formula
{
    using MathNet.Numerics;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using System.Net;

    /// <summary>
    /// <para>
    /// Evaluates Data Validation constraints.
    /// </para>
    /// <para>
    /// For performance reasons, this class keeps a cache of all previously retrieved <see cref="DataValidation"/> instances.
    /// Be sure to call <see cref="clearAllCachedValues()" /> if any workbook validation definitions are
    /// added, modified, or deleted.
    /// </para>
    /// <para>
    /// Changing cell values should be fine, as long as the corresponding <see cref="WorkbookEvaluator.clearAllCachedResultValues()" />
    /// is called as well.
    /// </para>
    /// </summary>
    public class DataValidationEvaluator
    {

        /// <summary>
        /// <para>
        /// Expensive to compute, so cache them as they are retrieved.
        /// </para>
        /// <para>
        /// Sheets don't implement equals, and since its an interface,
        /// there's no guarantee instances won't be recreated on the fly by some implementation.
        /// So we use sheet name.
        /// </para>
        /// </summary>
        private Dictionary<String, List<IDataValidation>> validations = new Dictionary<String, List<IDataValidation>>();

        private IWorkbook workbook;
        private WorkbookEvaluator workbookEvaluator;

        /// <summary>
        /// Use the same formula evaluation context used for other operations, so cell value
        /// changes are automatically noticed
        /// </summary>
        /// <param name="wb">the workbook this operates on</param>
        /// <param name="provider">provider for formula evaluation</param>
        public DataValidationEvaluator(IWorkbook wb, IWorkbookEvaluatorProvider provider)
        {
            this.workbook = wb;
            this.workbookEvaluator = provider.GetWorkbookEvaluator();
        }

        /// <summary>
        /// </summary>
        /// <returns>evaluator</returns>
        protected WorkbookEvaluator GetWorkbookEvaluator()
        {
            return workbookEvaluator;
        }

        /// <summary>
        /// Call this whenever validation structures change,
        /// so future results stay in sync with the Workbook state.
        /// </summary>
        public void clearAllCachedValues()
        {
            validations.Clear();
        }

        /// <summary>
        /// Lazy load validations by sheet, since reading the CT* types is expensive
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns>The <see cref="DataValidation"/>s for the sheet</returns>
        private List<IDataValidation> GetValidations(ISheet sheet)
        {
            validations.TryGetValue(sheet.SheetName, out List<IDataValidation> dvs);
            if(dvs == null && !validations.ContainsKey(sheet.SheetName))
            {
                dvs = sheet.GetDataValidations();
                validations.Add(sheet.SheetName, dvs);
            }
            return dvs;
        }

        /// <summary>
        /// Finds and returns the <see cref="DataValidation"/> for the cell, if there is
        /// one. Lookup is based on the first match from
        /// <see cref="DataValidation.Regions" /> for the cell's sheet. DataValidation
        /// regions must be in the same sheet as the DataValidation. Allowed values
        /// expressions may reference other sheets, however.
        /// </summary>
        /// <param name="cell">reference to check - use this in case the cell does not actually exist yet</param>
        /// <returns>the DataValidation applicable to the given cell, or null if no
        /// validation applies
        /// </returns>
        public IDataValidation GetValidationForCell(CellReference cell)
        {
            DataValidationContext vc = GetValidationContextForCell(cell);
            return vc == null ? null : vc.Validation;
        }

        /// <summary>
        /// Finds and returns the <see cref="DataValidationContext"/> for the cell, if there is
        /// one. Lookup is based on the first match from
        /// <see cref="DataValidation.Regions" /> for the cell's sheet. DataValidation
        /// regions must be in the same sheet as the DataValidation. Allowed values
        /// expressions may reference other sheets, however.
        /// </summary>
        /// <param name="cell">reference to check</param>
        /// <returns>the DataValidationContext applicable to the given cell, or null if no
        /// validation applies
        /// </returns>
        public DataValidationContext GetValidationContextForCell(CellReference cell)
        {
            ISheet sheet = workbook.GetSheet(cell.SheetName);
            if(sheet == null)
                return null;
            List<IDataValidation> dataValidations = GetValidations(sheet);
            if(dataValidations == null)
                return null;
            foreach(IDataValidation dv in dataValidations)
            {
                CellRangeAddressList regions = dv.Regions;
                if(regions == null)
                    return null;
                // current implementation can't return null
                foreach(CellRangeAddressBase range in regions.CellRangeAddresses)
                {
                    if(range.IsInRange(cell))
                    {
                        return new DataValidationContext(dv, this, range, cell);
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// <para>
        /// If <see cref="getValidationForCell(CellReference)" /> returns an instance, and the
        /// <see cref="ValidationType"/> is <see cref="ValidationType.LIST" />, return the valid
        /// values, whether they are from a static list or cell range.
        /// </para>
        /// <para>
        /// For all other validation types, or no validation at all, this method
        /// returns null.
        /// </para>
        /// <para>
        /// This method could throw an exception if the validation type is not LIST,
        /// but since this method is mostly useful in UI contexts, null seems the
        /// easier path.
        /// </para>
        /// </summary>
        /// <param name="cell">reference to check - use this in case the cell does not actually exist yet</param>
        /// <returns>returns an unmodifiable <see cref="List"/> of <see cref="ValueEval"/>s if applicable, or
        /// null
        /// </returns>
        public List<ValueEval> GetValidationValuesForCell(CellReference cell)
        {
            DataValidationContext context = GetValidationContextForCell(cell);

            if(context == null)
                return null;

            return GetValidationValuesForConstraint(context);
        }

        /// <summary>
        /// static so enums can reference it without creating a whole instance
        /// </summary>
        /// <returns>returns an unmodifiable <see cref="List"/> of <see cref="ValueEval"/>s, which may be empty</returns>
        protected static List<ValueEval> GetValidationValuesForConstraint(DataValidationContext context)
        {
            IDataValidationConstraint val = context.Validation.ValidationConstraint;
            if(val.GetValidationType() != ValidationType.LIST)
                return null;

            string formula = val.Formula1;

            List<ValueEval> values = new List<ValueEval>();

            if(val.ExplicitListValues != null && val.ExplicitListValues.Length > 0)
            {
                // assumes parsing interprets the overloaded property right for XSSF
                foreach(string s in val.ExplicitListValues)
                {
                    if(s != null)
                        values.Add(new StringEval(s)); // constructor throws exception on null
                }
            }
            else if(formula != null)
            {
                // evaluate formula for cell refs then Get their values
                // note this should return the raw formula result, not the "unwrapped" version that returns a single value.
                ValueEval eval = context.Evaluator.GetWorkbookEvaluator().EvaluateList(formula, context.Target, context.Region);
                // formula is a StringEval if the validation is by a fixed list.  Use the explicit list later.
                // there is no way from the model to tell if the list is fixed values or formula based.
                if(eval is TwoDEval)
                {
                    TwoDEval twod = (TwoDEval) eval;
                    for(int i = 0; i < twod.Height; i++)
                    {
                        ValueEval cellValue = twod.GetValue(i,  0);
                        values.Add(cellValue);
                    }
                }
            }
            return values;
        }

        /// <summary>
        /// <para>
        /// Use the validation returned by <see cref="getValidationForCell(CellReference)" /> if you
        /// want the error display details. This is the validation checked by this
        /// method, which attempts to replicate Excel's data validation rules.
        /// </para>
        /// <para>
        /// Note that to properly apply some validations, care must be taken to
        /// offset the base validation formula by the relative position of the
        /// current cell, or the wrong value is checked.
        /// </para>
        /// </summary>
        /// <param name="cellRef">The reference of the cell to evaluate</param>
        /// <returns>true if the cell has no validation or the cell value passes the
        /// defined validation, false if it Assert.Fails
        /// </returns>
        public bool IsValidCell(CellReference cellRef)
        {
            DataValidationContext context = GetValidationContextForCell(cellRef);

            if(context == null)
                return true;

            ICell cell = SheetUtil.GetCell(workbook.GetSheet(cellRef.SheetName), cellRef.Row, cellRef.Col);

            // now we can validate the cell

            // if empty, return not allowed flag
            if(cell == null
                || IsType(cell, CellType.Blank)
                || (IsType(cell, CellType.String)
                    && (cell.StringCellValue == null || string.IsNullOrEmpty(cell.StringCellValue))
                   )
               )
            {
                return context.Validation.EmptyCellAllowed;
            }

            // cell has a value

            return ValidationEnum.IsValid(cell, context);
        }

        /// <summary>
        /// Note that this assumes the cell cached value is up to date and in sync with data edits
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="type"></param>
        /// <returns>true if the cell or cached cell formula result type match the given type</returns>
        public static bool IsType(ICell cell, CellType type)
        {
            CellType cellType = cell.CellType;
            return cellType == type
                  || (cellType == CellType.Formula
                      && cell.CachedFormulaResultType == type
                     );
        }

        public class ValidationEnum
        {
            private static readonly ValidationEnum ANY = new ValidationEnum_ANY();
            private static readonly ValidationEnum INTEGER = new ValidationEnum_INTEGER();
            private static readonly ValidationEnum DECIMAL = new ValidationEnum_DECIMAL();
            private static readonly ValidationEnum LIST = new ValidationEnum_LIST();
            private static readonly ValidationEnum DATE = new ValidationEnum_DATE();
            private static readonly ValidationEnum TIME = new ValidationEnum_TIME();
            private static readonly ValidationEnum TEXT_LENGTH = new ValidationEnum_TEXT_LENGTH();
            private static readonly ValidationEnum FORMULA = new ValidationEnum_FORMULA();
            public static Dictionary<int, ValidationEnum> Values = new Dictionary<int, ValidationEnum>()
            {
                { ValidationType.ANY, ANY },
                { ValidationType.INTEGER, INTEGER },
                { ValidationType.DECIMAL, DECIMAL },
                { ValidationType.LIST, LIST },
                { ValidationType.DATE, DATE },
                { ValidationType.TIME, TIME },
                { ValidationType.TEXT_LENGTH, TEXT_LENGTH },
                { ValidationType.FORMULA, FORMULA },
            };
            private class ValidationEnum_ANY : ValidationEnum
            {
                public override bool IsValidValue(ICell cell, DataValidationContext context)
                {
                    return true;
                }
            }

            private class ValidationEnum_INTEGER : ValidationEnum
            {
                public override bool IsValidValue(ICell cell, DataValidationContext context)
                {
                    if(base.IsValidValue(cell, context))
                    {
                        // we know it is a number in the proper range, now check if it is an int
                        double value = cell.NumericCellValue; // can't Get here without a valid numeric value
                        return value.CompareTo((int)value) == 0;
                    }
                    return false;
                }
            }
            private class ValidationEnum_DECIMAL : ValidationEnum { }
            private class ValidationEnum_LIST : ValidationEnum
            {
                public override bool IsValidValue(ICell cell, DataValidationContext context)
                {
                    List<ValueEval> valueList = GetValidationValuesForConstraint(context);
                    if(valueList == null)
                        return true; // special case

                    // compare cell value to each item
                    foreach(ValueEval listVal in valueList)
                    {
                        ValueEval comp = listVal is RefEval ? ((RefEval) listVal).GetInnerValueEval(context.SheetIndex) : listVal;

                        // any value is valid if the list contains a blank value per Excel help
                        if(comp is BlankEval)
                            return true;
                        if(comp is ErrorEval)
                            continue; // nothing to check
                        if(comp is BoolEval)
                        {
                            if(IsType(cell, CellType.Boolean) && ((BoolEval) comp).BooleanValue == cell.BooleanCellValue)
                            {
                                return true;
                            }
                            else
                            {
                                continue; // check the rest
                            }
                        }
                        if(comp is NumberEval)
                        {
                            // could this have trouble with double precision/rounding errors and date/time values?
                            // do we need to allow a "close enough" double fractional range?
                            // I see 17 digits After the decimal separator in XSSF files, and for time values,
                            // there are sometimes discrepancies in the final decimal place.  
                            // I don't have a validation test case yet though. - GW
                            if(IsType(cell, CellType.Numeric) && ((NumberEval) comp).NumberValue == cell.NumericCellValue)
                            {
                                return true;
                            }
                            else
                            {
                                continue; // check the rest
                            }
                        }
                        if(comp is StringEval)
                        {
                            // interestingly, in testing, a validation value of the string "TRUE" or "true" 
                            // did not match a bool cell value of TRUE - so apparently cell type matters
                            // also, Excel validation is case insensitive - "true" is valid for the list value "TRUE"
                            if(IsType(cell, CellType.String) && ((StringEval) comp).StringValue.Equals(cell.StringCellValue, StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                            else
                            {
                                continue; // check the rest;
                            }
                        }
                    }
                    return false; // no matches
                }
            }
            private class ValidationEnum_DATE : ValidationEnum { }
            private class ValidationEnum_TIME : ValidationEnum { }
            private class ValidationEnum_TEXT_LENGTH : ValidationEnum
            {
                public override bool IsValidValue(ICell cell, DataValidationContext context)
                {
                    if(!IsType(cell, CellType.String))
                        return false;
                    string v = cell.StringCellValue;
                    return IsValidNumericValue(v.Length, context);
                }
            }
            private class ValidationEnum_FORMULA : ValidationEnum
            {
                /**
                 * Note the formula result must either be a bool result, or anything not in error.
                 * If bool, value must be true to pass, anything else valid is also passing, errors Assert.Fail.
                 * @see NPOI.SS.Formula.DataValidationEvaluator.ValidationEnum#isValidValue(NPOI.SS.UserModel.Cell, NPOI.SS.UserModel.DataValidationConstraint, NPOI.SS.Formula.WorkbookEvaluator)
                 */
                public override bool IsValidValue(ICell cell, DataValidationContext context)
                {
                    // unwrapped single value
                    ValueEval comp = context.Evaluator.GetWorkbookEvaluator().Evaluate(context.Formula1, context.Target, context.Region);
                    if(comp is RefEval)
                    {
                        comp = ((RefEval) comp).GetInnerValueEval(((RefEval) comp).FirstSheetIndex);
                    }

                    if(comp is BlankEval)
                        return true;
                    if(comp is ErrorEval)
                        return false;
                    if(comp is BoolEval)
                    {
                        return ((BoolEval) comp).BooleanValue;
                    }
                    // empirically tested in Excel - 0=false, any other number = true/valid
                    // see test file DataValidationEvaluations.xlsx
                    if(comp is NumberEval)
                    {
                        return ((NumberEval) comp).NumberValue != 0;
                    }
                    return false; // anything else is false, such as text
                }
            }

            public virtual bool IsValidValue(ICell cell, DataValidationContext context)
            {
                return IsValidNumericCell(cell, context);
            }

            /**
             * Uses the cell value, which may be the cached formula result value.
             * We won't re-evaluate cells here.  This validation would be After the cell value was updated externally.
             * Excel allows invalid values through methods like copy/paste, and only validates them when the user 
             * interactively edits the cell.   
             * @return if the cell is a valid numeric cell for the validation or not
             */
            protected bool IsValidNumericCell(ICell cell, DataValidationContext context)
            {
                if(!IsType(cell, CellType.Numeric))
                    return false;

                Double value = cell.NumericCellValue;
                return IsValidNumericValue(value, context);
            }

            /**
             * Is the number a valid option for the validation?
             */
            protected bool IsValidNumericValue(Double value, DataValidationContext context)
            {
                try
                {
                    Double? t1 = EvalOrConstant(context.Formula1, context);
                    // per Excel, a blank value for a numeric validation constraint formula validates true
                    if(t1 == null)
                        return true;
                    Double? t2 = null;
                    if(context.Operator == OperatorType.BETWEEN || context.Operator == OperatorType.NOT_BETWEEN)
                    {
                        t2 = EvalOrConstant(context.Formula2, context);
                        // per Excel, a blank value for a numeric validation constraint formula validates true
                        if(t2 == null)
                            return true;
                    }
                    return OperatorEnum.Values[context.Operator].IsValid(value, t1.Value, t2.Value);
                }
                catch(FormatException e)
                {
                    // one or both formulas are in error, not evaluating to a number, so the validation is false per Excel's behavior.
                    return false;
                }
            }

            /**
             * Evaluate a numeric formula value as either a constant or numeric expression.
             * Note that Excel treats validations with constraint formulas that evaluate to null as valid,
             * but evaluations in error or non-numeric are marked invalid.
             * @param formula
             * @param context
             * @return numeric value or null if not defined or the formula evaluates to an empty/missing cell.
             * @throws NumberFormatException if the formula is non-numeric when it should be
             */
            private static Double? EvalOrConstant(string formula, DataValidationContext context)
            {
                if(formula == null || string.IsNullOrEmpty(formula.Trim()))
                    return null; // shouldn't happen, but just in case
                try
                {
                    return Double.Parse(formula);
                }
                catch(FormatException e)
                {
                    // must be an expression, then.  Overloading by Excel in the file formats.
                }
                // note the call to the "unwrapped" version, which returns a single value
                ValueEval eval = context.Evaluator.GetWorkbookEvaluator().Evaluate(formula, context.Target, context.Region);
                if(eval is RefEval)
                {
                    eval = ((RefEval) eval).GetInnerValueEval(((RefEval) eval).FirstSheetIndex);
                }
                if(eval is BlankEval)
                    return null;
                if(eval is NumberEval)
                    return ((NumberEval) eval).NumberValue;
                if(eval is StringEval)
                {
                    string value = ((StringEval) eval).StringValue;
                    if(value == null || string.IsNullOrEmpty(value.Trim()))
                        return null;
                    // try to parse the cell value as a double and return it 
                    return Double.Parse(value);
                }
                throw new FormatException("Formula '" + formula + "' evaluates to something other than a number");
            }
            /// <summary>
            /// Validates against the type defined in context, as an index of the enum values array.
            /// </summary>
            /// <param name="cell">Cell to check validity of</param>
            /// <param name="context">The Data Validation to check against</param>
            /// <returns>true if validation passes</returns>
            /// <exception cref="ArrayIndexOutOfBoundsException">if the constraint type is an invalid index</exception>
            public static bool IsValid(ICell cell, DataValidationContext context)
            {
                return Values[context.Validation.ValidationConstraint.GetValidationType()].IsValidValue(cell, context);
            }
        }

        public class OperatorEnum
        {
            public static readonly OperatorEnum BETWEEN = new Operator_BETWEEN();
            public static readonly OperatorEnum NOT_BETWEEN = new Operator_NOT_BETWEEN();
            public static readonly OperatorEnum EQUAL = new Operator_EQUAL();
            public static readonly OperatorEnum NOT_EQUAL = new Operator_NOT_EQUAL();
            public static readonly OperatorEnum GREATER_THAN = new Operator_GREATER_THAN();
            public static readonly OperatorEnum LESS_THAN = new Operator_LESS_THAN();
            public static readonly OperatorEnum GREATER_OR_EQUAL = new Operator_GREATER_OR_EQUAL();
            public static readonly OperatorEnum LESS_OR_EQUAL = new Operator_LESS_OR_EQUAL();

            public static readonly OperatorEnum IGNORED = BETWEEN;

            public static readonly Dictionary<int, OperatorEnum> Values = new Dictionary<int, OperatorEnum>()
            {
                { OperatorType.BETWEEN, BETWEEN },
                { OperatorType.NOT_BETWEEN, NOT_BETWEEN },
                { OperatorType.EQUAL, EQUAL },
                { OperatorType.NOT_EQUAL, NOT_EQUAL },
                { OperatorType.GREATER_THAN, GREATER_THAN },
                { OperatorType.LESS_THAN, LESS_THAN },
                { OperatorType.GREATER_OR_EQUAL, GREATER_OR_EQUAL },
                { OperatorType.LESS_OR_EQUAL, LESS_OR_EQUAL },
            };
            /**
             * Evaluates comparison using operator instance rules
             * @param cellValue won't be null, assumption is previous checks handled that
             * @param v1 if null, value assumed invalid, anything passes, per Excel behavior
             * @param v2 null if not needed.  If null when needed, assume anything passes, per Excel behavior
             * @return true if the comparison is valid
             */
            public virtual bool IsValid(Double cellValue, Double v1, Double v2)
            {
                throw new NotImplementedException();
            }

            private class Operator_BETWEEN : OperatorEnum
            {
                public override bool IsValid(Double cellValue, Double v1, Double v2)
                {
                    return cellValue.CompareTo(v1) >= 0 && cellValue.CompareTo(v2) <= 0;
                }
            }

            private class Operator_NOT_BETWEEN : OperatorEnum
            {
                public override bool IsValid(Double cellValue, Double v1, Double v2)
                {
                    return cellValue.CompareTo(v1) < 0 || cellValue.CompareTo(v2) > 0;
                }
            }

            private class Operator_EQUAL : OperatorEnum
            {
                public override bool IsValid(Double cellValue, Double v1, Double v2)
                {
                    return cellValue.CompareTo(v1) == 0;
                }
            }

            private class Operator_NOT_EQUAL : OperatorEnum
            {
                public override bool IsValid(Double cellValue, Double v1, Double v2)
                {
                    return cellValue.CompareTo(v1) != 0;
                }
            }
            private class Operator_GREATER_THAN : OperatorEnum
            {
                public override bool IsValid(Double cellValue, Double v1, Double v2)
                {
                    return cellValue.CompareTo(v1) > 0;
                }
            }
            private class Operator_LESS_THAN : OperatorEnum
            {
                public override bool IsValid(Double cellValue, Double v1, Double v2)
                {
                    return cellValue.CompareTo(v1) < 0;
                }
            }
            private class Operator_GREATER_OR_EQUAL : OperatorEnum
            {
                public override bool IsValid(Double cellValue, Double v1, Double v2)
                {
                    return cellValue.CompareTo(v1) >= 0;
                }
            }
            private class Operator_LESS_OR_EQUAL : OperatorEnum
            {
                public override bool IsValid(Double cellValue, Double v1, Double v2)
                {
                    return cellValue.CompareTo(v1) <= 0;
                }
            }
        }


        /**
         * This class organizes and encapsulates all the pieces of information related to a single
         * data validation configuration for a single cell.  It cleanly separates the validation region,
         * the cells it applies to, the specific cell this instance references, and the validation
         * configuration and current values if applicable.
         */
        public class DataValidationContext
        {
            private IDataValidation dv;
            private DataValidationEvaluator dve;
            private CellRangeAddressBase region;
            private CellReference target;

            /**
             *
             * @param dv
             * @param dve
             * @param region
             * @param target
             */
            public DataValidationContext(IDataValidation dv, DataValidationEvaluator dve, CellRangeAddressBase region, CellReference target)
            {
                this.dv = dv;
                this.dve = dve;
                this.region = region;
                this.target = target;
            }
            /**
             * @return the dv
             */
            public IDataValidation Validation
            {
                get
                {
                    return dv;
                }
            }
            /**
             * @return the dve
             */
            public DataValidationEvaluator Evaluator
            {
                get
                {
                    return dve;
                }
            }
            /**
             * @return the region
             */
            public CellRangeAddressBase Region
            {
                get
                {
                    return region;
                }
            }
            /**
             * @return the target
             */
            public CellReference Target
            {
                get
                {
                    return target;
                }
            }

            public int OffsetColumns
            {
                get
                {
                    return target.Col - region.FirstColumn;
                }
            }

            public int OffsetRows
            {
                get
                {
                    return target.Row - region.FirstRow;
                }
            }

            public int SheetIndex
            {
                get
                {
                    return dve.GetWorkbookEvaluator().GetSheetIndex(target.SheetName);
                }
            }

            public string Formula1
            {
                get
                {
                    return dv.ValidationConstraint.Formula1;
                }
            }

            public string Formula2
            {
                get
                {
                    return dv.ValidationConstraint.Formula2;
                }
            }

            public int Operator
            {
                get
                {
                    return dv.ValidationConstraint.Operator;
                }

            }

        }
    }

}
