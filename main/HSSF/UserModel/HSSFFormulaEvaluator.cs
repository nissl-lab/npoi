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
    using NPOI.SS.Formula.UDF;
    using NPOI.SS.UserModel;
    using System.Collections.Generic;

    /**
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * 
     */
    public class HSSFFormulaEvaluator : BaseFormulaEvaluator
    {
        // params to lookup the right constructor using reflection
        private static readonly Type[] VALUE_CONTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg) };

        private static readonly Type[] AREA3D_CONSTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg), typeof(ValueEval[]) };

        private static readonly Type[] REFERENCE_CONSTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg), typeof(ValueEval) };

        private static readonly Type[] REF3D_CONSTRUCTOR_CLASS_ARRAY = new Type[] { typeof(Ptg), typeof(ValueEval) };

        // Maps for mapping *Eval to *Ptg
        private static readonly Hashtable VALUE_EVALS_MAP = new Hashtable();

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
        protected IWorkbook _book;

        public HSSFFormulaEvaluator(IWorkbook workbook)
            : this(workbook, null)
        {
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
            : base(new WorkbookEvaluator(HSSFEvaluationWorkbook.Create(workbook), stabilityClassifier, udfFinder))
        {
            _book = workbook;
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

        protected override IRichTextString CreateRichTextString(string str)
        {
            return new HSSFRichTextString(str);
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
            BaseFormulaEvaluator.SetupEnvironment(workbookNames, evaluators);
        }

        public override void SetupReferencedWorkbooks(Dictionary<String, IFormulaEvaluator> evaluators)
        {
            CollaboratingWorkbooksEnvironment.SetupFormulaEvaluator(evaluators);
        }

        /**
         * Should be called to tell the cell value cache that the specified (value or formula) cell 
         * has changed.
         * Failure to call this method after changing cell values will cause incorrect behaviour
         * of the evaluate~ methods of this class
         */
        public override void NotifyUpdateCell(ICell cell)
        {
            _bookEvaluator.NotifyUpdateCell(new HSSFEvaluationCell(cell));
        }
        /**
         * Should be called to tell the cell value cache that the specified cell has just been
         * deleted. 
         * Failure to call this method after changing cell values will cause incorrect behaviour
         * of the evaluate~ methods of this class
         */
        public override void NotifyDeleteCell(ICell cell)
        {
            _bookEvaluator.NotifyDeleteCell(new HSSFEvaluationCell(cell));
        }

        /**
         * Should be called to tell the cell value cache that the specified (value or formula) cell
         * has changed.
         * Failure to call this method after changing cell values will cause incorrect behaviour
         * of the evaluate~ methods of this class
         */
        public override void NotifySetFormula(ICell cell)
        {
            _bookEvaluator.NotifyUpdateCell(new HSSFEvaluationCell(cell));
        }

        /**
         * Returns a CellValue wrapper around the supplied ValueEval instance.
         * @param cell
         */
        protected override CellValue EvaluateFormulaCellValue(ICell cell)
        {
            ValueEval eval = _bookEvaluator.Evaluate(new HSSFEvaluationCell((HSSFCell)cell));
            if (eval is BoolEval be)
            {
                return CellValue.ValueOf(be.BooleanValue);
            }
            if (eval is NumberEval numberEval)
            {
                return new CellValue(numberEval.NumberValue);
            }
            if (eval is StringEval ne)
            {
                return new CellValue(ne.StringValue);
            }
            if (eval is ErrorEval errorEval)
            {
                return CellValue.GetError(errorEval.ErrorCode);
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
        public override ICell EvaluateInCell(ICell cell)
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
        public new static void EvaluateAllFormulaCells(IWorkbook wb)
        {
            BaseFormulaEvaluator.EvaluateAllFormulaCells(wb);
        }

        public override void EvaluateAll()
        {
            HSSFFormulaEvaluator.EvaluateAllFormulaCells(_book, this);
        }

    }
}