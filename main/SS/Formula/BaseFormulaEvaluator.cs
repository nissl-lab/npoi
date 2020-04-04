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

namespace NPOI.SS.Formula
{
    using System;
    using System.Collections.Generic;
    using NPOI.SS.UserModel;
    using NPOI.Util;

    /**
    * Common functionality across file formats for Evaluating formula cells.<p/>
    */
    public abstract class BaseFormulaEvaluator : IFormulaEvaluator, IWorkbookEvaluatorProvider
    {
        protected WorkbookEvaluator _bookEvaluator;

        protected BaseFormulaEvaluator(WorkbookEvaluator bookEvaluator)
        {
            this._bookEvaluator = bookEvaluator;
        }

        /**
         * Coordinates several formula Evaluators together so that formulas that involve external
         * references can be Evaluated.
         * @param workbookNames the simple file names used to identify the workbooks in formulas
         * with external links (for example "MyData.xls" as used in a formula "[MyData.xls]Sheet1!A1")
         * @param Evaluators all Evaluators for the full Set of workbooks required by the formulas.
         */
        public static void SetupEnvironment(String[] workbookNames, BaseFormulaEvaluator[] Evaluators)
        {
            WorkbookEvaluator[] wbEvals = new WorkbookEvaluator[Evaluators.Length];
            for (int i = 0; i < wbEvals.Length; i++)
            {
                wbEvals[i] = Evaluators[i]._bookEvaluator;
            }
            CollaboratingWorkbooksEnvironment.Setup(workbookNames, wbEvals);
        }


        public virtual void SetupReferencedWorkbooks(Dictionary<String, IFormulaEvaluator> evaluators)
        {
            CollaboratingWorkbooksEnvironment.SetupFormulaEvaluator(evaluators);
        }


        public WorkbookEvaluator GetWorkbookEvaluator()
        {
            return _bookEvaluator;
        }

        /**
         * Should be called whenever there are major Changes (e.g. moving sheets) to input cells
         * in the Evaluated workbook.  If performance is not critical, a single call to this method
         * may be used instead of many specific calls to the Notify~ methods.
         *
         * Failure to call this method After changing cell values will cause incorrect behaviour
         * of the Evaluate~ methods of this class
         */

        public void ClearAllCachedResultValues()
        {
            _bookEvaluator.ClearAllCachedResultValues();
        }

        /**
         * If cell Contains a formula, the formula is Evaluated and returned,
         * else the CellValue simply copies the appropriate cell value from
         * the cell and also its cell type. This method should be preferred over
         * EvaluateInCell() when the call should not modify the contents of the
         * original cell.
         *
         * @param cell may be <code>null</code> signifying that the cell is not present (or blank)
         * @return <code>null</code> if the supplied cell is <code>null</code> or blank
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
                default:
                    throw new InvalidOperationException("Bad cell type (" + cell.CellType + ")");
            }
        }

        /**
         * If cell Contains formula, it Evaluates the formula, and
         *  Puts the formula result back into the cell, in place
         *  of the old formula.
         * Else if cell does not contain formula, this method leaves
         *  the cell unChanged.
         * Note that the same instance of HSSFCell is returned to
         * allow chained calls like:
         * <pre>
         * int EvaluatedCellType = Evaluator.EvaluateInCell(cell).CellType;
         * </pre>
         * Be aware that your cell value will be Changed to hold the
         *  result of the formula. If you simply want the formula
         *  value computed for you, use {@link #EvaluateFormulaCellEnum(Cell)}}
         * @param cell
         * @return the {@code cell} that was passed in, allowing for chained calls
         */

        public virtual ICell EvaluateInCell(ICell cell)
        {
            if (cell == null)
            {
                return null;
            }
            ICell result = cell;
            if (cell.CellType == CellType.Formula)
            {
                CellValue cv = EvaluateFormulaCellValue(cell);
                SetCellValue(cell, cv);
                SetCellType(cell, cv); // cell will no longer be a formula cell
            }
            return result;
        }

        protected abstract CellValue EvaluateFormulaCellValue(ICell cell);

        /**
         * If cell Contains formula, it Evaluates the formula, and saves the result of the formula. The
         * cell remains as a formula cell. If the cell does not contain formula, this method returns -1
         * and leaves the cell unChanged.
         *
         * Note that the type of the <em>formula result</em> is returned, so you know what kind of
         * cached formula result is also stored with  the formula.
         * <pre>
         * int EvaluatedCellType = Evaluator.EvaluateFormulaCell(cell);
         * </pre>
         * Be aware that your cell will hold both the formula, and the result. If you want the cell
         * Replaced with the result of the formula, use {@link #EvaluateInCell(NPOI.SS.UserModel.Cell)}
         * @param cell The cell to Evaluate
         * @return -1 for non-formula cells, or the type of the <em>formula result</em>
         */
        public CellType EvaluateFormulaCell(ICell cell)
        {
            return EvaluateFormulaCellEnum(cell);
        }

        /**
         * If cell Contains formula, it Evaluates the formula,
         *  and saves the result of the formula. The cell
         *  remains as a formula cell.
         * Else if cell does not contain formula, this method leaves
         *  the cell unChanged.
         * Note that the type of the formula result is returned,
         *  so you know what kind of value is also stored with
         *  the formula.
         * <pre>
         * ICellType EvaluatedCellType = Evaluator.EvaluateFormulaCellEnum(cell);
         * </pre>
         * Be aware that your cell will hold both the formula,
         *  and the result. If you want the cell Replaced with
         *  the result of the formula, use {@link #Evaluate(NPOI.SS.UserModel.Cell)} }
         * @param cell The cell to Evaluate
         * @return The type of the formula result (the cell's type remains as CellType.FORMULA however)
         *         If cell is not a formula cell, returns {@link CellType#_NONE} rather than throwing an exception.
         * @since POI 3.15 beta 3
         */
        public virtual CellType EvaluateFormulaCellEnum(ICell cell)
        {
            if (cell == null || cell.CellType != CellType.Formula)
            {
                return CellType.Unknown;
            }
            CellValue cv = EvaluateFormulaCellValue(cell);
            // cell remains a formula cell, but the cached value is Changed
            SetCellValue(cell, cv);
            return cv.CellType;
        }

        protected static void SetCellType(ICell cell, CellValue cv)
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
                    // never happens - blanks eventually Get translated to zero
                    throw new ArgumentException("This should never happen. Blanks eventually Get translated to zero.");
                case CellType.Formula:
                    // this will never happen, we have already Evaluated the formula
                    throw new ArgumentException("This should never happen. Formulas should have already been Evaluated.");
                default:
                    throw new InvalidOperationException("Unexpected cell value type (" + cellType + ")");
            }
        }

        protected abstract IRichTextString CreateRichTextString(String str);

        protected void SetCellValue(ICell cell, CellValue cv)
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
                    cell.SetCellValue(CreateRichTextString(cv.StringValue));
                    break;
                case CellType.Blank:
                // never happens - blanks eventually Get translated to zero
                case CellType.Formula:
                // this will never happen, we have already Evaluated the formula
                default:
                    throw new InvalidOperationException("Unexpected cell value type (" + cellType + ")");
            }
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
        public static void EvaluateAllFormulaCells(IWorkbook wb)
        {
            IFormulaEvaluator evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();
            EvaluateAllFormulaCells(wb, evaluator);
        }
        protected static void EvaluateAllFormulaCells(IWorkbook wb, IFormulaEvaluator evaluator)
        {
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);

                foreach (IRow r in sheet)
                {
                    foreach (ICell c in r)
                    {
                        if (c.CellType == CellType.Formula)
                        {
                            evaluator.EvaluateFormulaCell(c);
                        }
                    }
                }
            }
        }

        public abstract void NotifySetFormula(ICell cell);

        public abstract void NotifyDeleteCell(ICell cell);

        public abstract void NotifyUpdateCell(ICell cell);

        public abstract void EvaluateAll();

        /** {@inheritDoc} */

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