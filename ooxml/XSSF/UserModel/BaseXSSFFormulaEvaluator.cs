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

namespace NPOI.XSSF.UserModel
{
    using System;
    using System.Collections.Generic;
    using NPOI.SS.Formula;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;

    /**
     * Internal POI use only - parent of XSSF and SXSSF formula Evaluators
     */
    public abstract class BaseXSSFFormulaEvaluator : BaseFormulaEvaluator
    {

        protected BaseXSSFFormulaEvaluator(WorkbookEvaluator bookEvaluator)
            : base(bookEvaluator)
        {
        }

        public override void NotifySetFormula(ICell cell)
        {
            _bookEvaluator.NotifyUpdateCell(new XSSFEvaluationCell((XSSFCell)cell));
        }
        public override void NotifyDeleteCell(ICell cell)
        {
            _bookEvaluator.NotifyDeleteCell(new XSSFEvaluationCell((XSSFCell)cell));
        }
        public override void NotifyUpdateCell(ICell cell)
        {
            _bookEvaluator.NotifyUpdateCell(new XSSFEvaluationCell((XSSFCell)cell));
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
         * int EvaluatedCellType = Evaluator.EvaluateFormulaCell(cell);
         * </pre>
         * Be aware that your cell will hold both the formula,
         *  and the result. If you want the cell Replaced with
         *  the result of the formula, use {@link #Evaluate(NPOI.SS.UserModel.Cell)} }
         * @param cell The cell to Evaluate
         * @return The type of the formula result (the cell's type remains as HSSFCellType.FORMULA however)
         */
        [Obsolete("deprecated POI 3.15 beta 3. Will be deleted when we make the CellType enum transition. See bug 59791.")]
        public override CellType EvaluateFormulaCellEnum(ICell cell)
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

        /**
         * If cell Contains formula, it Evaluates the formula, and
         *  Puts the formula result back into the cell, in place
         *  of the old formula.
         * Else if cell does not contain formula, this method leaves
         *  the cell unChanged.
         */
        protected void DoEvaluateInCell(ICell cell)
        {
            if (cell == null) return;
            if (cell.CellType == CellType.Formula)
            {
                CellValue cv = EvaluateFormulaCellValue(cell);
                SetCellType(cell, cv); // cell will no longer be a formula cell
                SetCellValue(cell, cv);
            }
        }

        private new static void SetCellValue(ICell cell, CellValue cv)
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
                    cell.SetCellValue(new XSSFRichTextString(cv.StringValue));
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
         * Turns a XSSFCell / SXSSFCell into a XSSFEvaluationCell
         */
        protected abstract IEvaluationCell ToEvaluationCell(ICell cell);

        /**
         * Returns a CellValue wrapper around the supplied ValueEval instance.
         */
        protected override CellValue EvaluateFormulaCellValue(ICell cell)
        {
            IEvaluationCell evalCell = ToEvaluationCell(cell);
            ValueEval eval = _bookEvaluator.Evaluate(evalCell);
            if (eval is NumberEval)
            {
                NumberEval ne = (NumberEval)eval;
                return new CellValue(ne.NumberValue);
            }
            if (eval is BoolEval)
            {
                BoolEval be = (BoolEval)eval;
                return CellValue.ValueOf(be.BooleanValue);
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
            throw new Exception("Unexpected eval class (" + eval.GetType().Name + ")");
        }
    }

}