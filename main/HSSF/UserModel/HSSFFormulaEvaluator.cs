/*
* Licensed to the Apache Software Foundation (ASF) Under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You Under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed Under the License is distributed on an "AS Is" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations Under the License.
*/

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Collections;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.PTG;
    using NPOI.SS.Formula.Udf;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * 
     */
    public class HSSFFormulaEvaluator : IFormulaEvaluator, IWorkbookEvaluatorProvider
    {
        private WorkbookEvaluator _bookEvaluator;
        // params to lookup the right constructor using reflection
        private static Type[] VALUE_CONTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg) };

        private static Type[] AREA3D_CONSTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg), typeof(ValueEval[]) };

        private static Type[] REFERENCE_CONSTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg), typeof(ValueEval) };

        private static Type[] REF3D_CONSTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg), typeof(ValueEval) };

        // Maps for mapping *Eval to *Ptg
        private static Hashtable VALUE_EVALS_MAP = new Hashtable();

        /*
         * Following is the mapping between the Ptg tokens returned 
         * by the FormulaParser and the *Eval classes that are used 
         * by the FormulaEvaluator
         */
        static HSSFFormulaEvaluator()
        {
            VALUE_EVALS_MAP[typeof(BoolPtg)] = typeof(BoolEval);
            VALUE_EVALS_MAP[typeof(IntPtg)] = typeof(NumberEval);
            VALUE_EVALS_MAP[typeof(NumberPtg)] = typeof(NumberEval);
            VALUE_EVALS_MAP[typeof(StringPtg)] = typeof(StringEval);

        }


        protected IRow row;
        protected ISheet sheet;
        protected IWorkbook workbook;
        [Obsolete]
        public HSSFFormulaEvaluator(ISheet sheet, IWorkbook workbook)
            : this(workbook)
        {
            this.sheet = sheet;
            this.workbook = workbook;
        }

        public HSSFFormulaEvaluator(IWorkbook workbook)
            : this(workbook, null)
        {
            this.workbook = workbook;
        }
        /**
         * @param stabilityClassifier used to optimise caching performance. Pass <code>null</code>
         * for the (conservative) assumption that any cell may have its definition changed after
         * evaluation begins.
         */
        public HSSFFormulaEvaluator(IWorkbook workbook, IStabilityClassifier stabilityClassifier)
            : this(workbook, stabilityClassifier, null)
        {

        }



        /**
         * @param udfFinder pass <code>null</code> for default (AnalysisToolPak only)
         */
        public HSSFFormulaEvaluator(IWorkbook workbook, IStabilityClassifier stabilityClassifier, UDFFinder udfFinder)
        {
            _bookEvaluator = new WorkbookEvaluator(HSSFEvaluationWorkbook.Create(workbook), stabilityClassifier, udfFinder);
        }

        /**
     * @param stabilityClassifier used to optimise caching performance. Pass <code>null</code>
     * for the (conservative) assumption that any cell may have its definition changed after
     * evaluation begins.
     * @param udfFinder pass <code>null</code> for default (AnalysisToolPak only)
     */
        public static HSSFFormulaEvaluator Create(IWorkbook workbook, IStabilityClassifier stabilityClassifier, UDFFinder udfFinder)
        {
            return new HSSFFormulaEvaluator(workbook, stabilityClassifier, udfFinder);
        }


        private static void SetCellType(ICell cell, CellValue cv)
        {
            CellType cellType = cv.CellType;
            switch (cellType)
            {
                case CellType.Boolean:
                case CellType.Error:
                case CellType.Numeric:
                case CellType.String:
                    cell.SetCellType(cellType);
                    return;
                case CellType.Blank:
                    // never happens - blanks eventually get translated to zero
                    break;
                case CellType.Formula:
                    // this will never happen, we have already evaluated the formula
                    break;
            }
            throw new InvalidOperationException("Unexpected cell value type (" + cellType + ")");
        }
        private static void SetCellValue(ICell cell, CellValue cv)
        {
            CellType cellType = cv.CellType;
            switch (cellType)
            {
                case CellType.Boolean:
                    cell.SetCellValue(cv.BooleanValue);
                    break;
                case CellType.Error:
                    cell.SetCellErrorValue((byte)cv.ErrorValue);
                    break;
                case CellType.Numeric:
                    cell.SetCellValue(cv.NumberValue);
                    break;
                case CellType.String:
                    cell.SetCellValue(new HSSFRichTextString(cv.StringValue));
                    break;
                //case CellType.Blank:
                //// never happens - blanks eventually get translated to zero
                //case CellType.Formula:
                //// this will never happen, we have already evaluated the formula
                default:
                    throw new InvalidOperationException("Unexpected cell value type (" + cellType + ")");
            }
        }
        /**
         * Coordinates several formula evaluators together so that formulas that involve external
         * references can be evaluated.
         * @param workbookNames the simple file names used to identify the workbooks in formulas
         * with external links (for example "MyData.xls" as used in a formula "[MyData.xls]Sheet1!A1")
         * @param evaluators all evaluators for the full set of workbooks required by the formulas. 
         */
        public static void SetupEnvironment(String[] workbookNames, HSSFFormulaEvaluator[] evaluators)
        {
            WorkbookEvaluator[] wbEvals = new WorkbookEvaluator[evaluators.Length];
            for (int i = 0; i < wbEvals.Length; i++)
            {
                wbEvals[i] = evaluators[i]._bookEvaluator;
            }
            CollaboratingWorkbooksEnvironment.Setup(workbookNames, wbEvals);
        }

        public void SetupReferencedWorkbooks(Dictionary<String, IFormulaEvaluator> evaluators)
        {
            CollaboratingWorkbooksEnvironment.SetupFormulaEvaluator(evaluators);
        }

        public WorkbookEvaluator GetWorkbookEvaluator()
        {
            return _bookEvaluator;
        }
        /**
         * If cell Contains a formula, the formula is Evaluated and returned,
         * else the CellValue simply copies the appropriate cell value from
         * the cell and also its cell type. This method should be preferred over
         * EvaluateInCell() when the call should not modify the contents of the
         * original cell. 
         * @param cell
         */
        /**
         * If cell contains a formula, the formula is evaluated and returned,
         * else the CellValue simply copies the appropriate cell value from
         * the cell and also its cell type. This method should be preferred over
         * evaluateInCell() when the call should not modify the contents of the
         * original cell.
         * 
         * @param cell may be <c>null</c> signifying that the cell is not present (or blank)
         * @return <c>null</c> if the supplied cell is <c>null</c> or blank
         */
        public CellValue Evaluate(ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return CellValue.ValueOf(cell.BooleanCellValue);
                case CellType.Error:
                    return CellValue.GetError(cell.ErrorCellValue);
                case CellType.Formula:
                    return EvaluateFormulaCellValue(cell);
                case CellType.Numeric:
                    return new CellValue(cell.NumericCellValue);
                case CellType.String:
                    return new CellValue(cell.RichStringCellValue.String);
                case CellType.Blank:
                    return null;
            }
            throw new InvalidOperationException("Bad cell type (" + cell.CellType + ")");
        }


        /**
         * Should be called whenever there are major changes (e.g. moving sheets) to input cells
         * in the evaluated workbook.  If performance is not critical, a single call to this method
         * may be used instead of many specific calls to the notify~ methods.
         *  
         * Failure to call this method after changing cell values will cause incorrect behaviour
         * of the evaluate~ methods of this class
         */
        public void ClearAllCachedResultValues()
        {
            _bookEvaluator.ClearAllCachedResultValues();
        }
        /**
         * Should be called to tell the cell value cache that the specified (value or formula) cell 
         * has changed.
         * Failure to call this method after changing cell values will cause incorrect behaviour
         * of the evaluate~ methods of this class
         */
        public void NotifyUpdateCell(ICell cell)
        {
            _bookEvaluator.NotifyUpdateCell(new HSSFEvaluationCell(cell));
        }
        /**
         * Should be called to tell the cell value cache that the specified cell has just been
         * deleted. 
         * Failure to call this method after changing cell values will cause incorrect behaviour
         * of the evaluate~ methods of this class
         */
        public void NotifyDeleteCell(ICell cell)
        {
            _bookEvaluator.NotifyDeleteCell(new HSSFEvaluationCell(cell));
        }

        /**
         * Should be called to tell the cell value cache that the specified (value or formula) cell
         * has changed.
         * Failure to call this method after changing cell values will cause incorrect behaviour
         * of the evaluate~ methods of this class
         */
        public void NotifySetFormula(ICell cell)
        {
            _bookEvaluator.NotifyUpdateCell(new HSSFEvaluationCell(cell));
        }

        /**
         * If cell Contains formula, it Evaluates the formula,
         *  and saves the result of the formula. The cell
         *  remains as a formula cell.
         * Else if cell does not contain formula, this method leaves
         *  the cell UnChanged. 
         * Note that the type of the formula result is returned,
         *  so you know what kind of value is also stored with
         *  the formula. 
         * <pre>
         * int EvaluatedCellType = evaluator.EvaluateFormulaCell(cell);
         * </pre>
         * Be aware that your cell will hold both the formula,
         *  and the result. If you want the cell Replaced with
         *  the result of the formula, use {@link #EvaluateInCell(HSSFCell)}
         * @param cell The cell to Evaluate
         * @return The type of the formula result (the cell's type remains as CellType.Formula however)
         */
        public CellType EvaluateFormulaCell(ICell cell)
        {
            if (cell == null || cell.CellType != CellType.Formula)
            {
                return CellType.Unknown;
            }
            CellValue cv = EvaluateFormulaCellValue(cell);
            // cell remains a formula cell, but the cached value is changed
            SetCellValue(cell, cv);
            return cv.CellType;
        }
        /**
         * Returns a CellValue wrapper around the supplied ValueEval instance.
         * @param cell
         */
        private CellValue EvaluateFormulaCellValue(ICell cell)
        {
            ValueEval eval = _bookEvaluator.Evaluate(new HSSFEvaluationCell((HSSFCell)cell));
            if (eval is BoolEval)
            {
                BoolEval be = (BoolEval)eval;
                return CellValue.ValueOf(be.BooleanValue);
            }
            if (eval is NumberEval)
            {
                NumberEval ne = (NumberEval)eval;
                return new CellValue(ne.NumberValue);
            }
            if (eval is StringEval)
            {
                StringEval ne = (StringEval)eval;
                return new CellValue(ne.StringValue);
            }
            if (eval is ErrorEval)
            {
                return CellValue.GetError(((ErrorEval)eval).ErrorCode);
            }
            throw new InvalidOperationException("Unexpected eval class (" + eval.GetType().Name + ")");
        }
        /**
         * If cell Contains formula, it Evaluates the formula, and
         *  puts the formula result back into the cell, in place
         *  of the old formula.
         * Else if cell does not contain formula, this method leaves
         *  the cell UnChanged. 
         * Note that the same instance of Cell is returned to 
         * allow chained calls like:
         * <pre>
         * int EvaluatedCellType = evaluator.EvaluateInCell(cell).CellType;
         * </pre>
         * Be aware that your cell value will be Changed to hold the
         *  result of the formula. If you simply want the formula
         *  value computed for you, use {@link #EvaluateFormulaCell(HSSFCell)}
         * @param cell
         */
        public ICell EvaluateInCell(ICell cell)
        {
            if (cell == null)
            {
                return null;
            }
            if (cell.CellType == CellType.Formula)
            {
                CellValue cv = EvaluateFormulaCellValue(cell);
                SetCellValue(cell, cv);
                SetCellType(cell, cv); // cell will no longer be a formula cell
            }
            return cell;
        }

        /**
         * Loops over all cells in all sheets of the supplied
         *  workbook.
         * For cells that contain formulas, their formulas are
         *  Evaluated, and the results are saved. These cells
         *  remain as formula cells.
         * For cells that do not contain formulas, no Changes
         *  are made.
         * This is a helpful wrapper around looping over all 
         *  cells, and calling EvaluateFormulaCell on each one.
         */
        public static void EvaluateAllFormulaCells(HSSFWorkbook wb)
        {
            EvaluateAllFormulaCells(wb, new HSSFFormulaEvaluator(wb));
        }
        /**
           * Loops over all cells in all sheets of the supplied
           *  workbook.
           * For cells that contain formulas, their formulas are
           *  evaluated, and the results are saved. These cells
           *  remain as formula cells.
           * For cells that do not contain formulas, no changes
           *  are made.
           * This is a helpful wrapper around looping over all
           *  cells, and calling evaluateFormulaCell on each one.
           */
        public static void EvaluateAllFormulaCells(IWorkbook wb)
        {
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            EvaluateAllFormulaCells(wb, evaluator);
        }
        private static void EvaluateAllFormulaCells(IWorkbook wb, IFormulaEvaluator evaluator)
        {
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);

                for (IEnumerator it = sheet.GetRowEnumerator(); it.MoveNext(); )
                {
                    IRow r = (IRow)it.Current;
                    foreach (ICell c in r.Cells)
                    {
                        if (c.CellType == CellType.Formula)
                        {
                            evaluator.EvaluateFormulaCell(c);
                        }
                    }
                }
            }
        }


        public void EvaluateAll()
        {
            HSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook, this);
        }

        /**
         * Whether to ignore missing references to external workbooks and
         * use cached formula results in the main workbook instead.
         * <p>
         * In some cases exetrnal workbooks referenced by formulas in the main workbook are not avaiable.
         * With this method you can control how POI handles such missing references:
         * <ul>
         *     <li>by default ignoreMissingWorkbooks=false and POI throws {@link org.apache.poi.ss.formula.CollaboratingWorkbooksEnvironment.WorkbookNotFoundException}
         *     if an external reference cannot be resolved</li>
         *     <li>if ignoreMissingWorkbooks=true then POI uses cached formula result
         *     that already exists in the main workbook</li>
         * </ul>
         *</p>
         * @param ignore whether to ignore missing references to external workbooks
         */
        public bool IgnoreMissingWorkbooks
        {
            get { return _bookEvaluator.IgnoreMissingWorkbooks; }
            set { _bookEvaluator.IgnoreMissingWorkbooks = value; }
        }

        /** {@inheritDoc} */
        public bool DebugEvaluationOutputForNextEval
        {
            get { return _bookEvaluator.DebugEvaluationOutputForNextEval; }
            set { _bookEvaluator.DebugEvaluationOutputForNextEval = value; }
            
        }

    }
}