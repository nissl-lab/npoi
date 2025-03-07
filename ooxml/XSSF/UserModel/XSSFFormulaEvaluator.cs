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

using NPOI.SS.Formula;
using NPOI.XSSF.UserModel;
using System;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.UDF;
using System.Collections.Generic;
namespace NPOI.XSSF.UserModel
{

    /**
     * Evaluates formula cells.<p/>
     *
     * For performance reasons, this class keeps a cache of all previously calculated intermediate
     * cell values.  Be sure to call {@link #ClearAllCachedResultValues()} if any workbook cells are Changed between
     * calls to Evaluate~ methods on this class.
     *
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     * @author Josh Micich
     */
    public class XSSFFormulaEvaluator : BaseXSSFFormulaEvaluator
    {

        private readonly XSSFWorkbook _book;

        public XSSFFormulaEvaluator(IWorkbook workbook)
            : this(workbook as XSSFWorkbook, null, null)
        { }
        public XSSFFormulaEvaluator(XSSFWorkbook workbook)
            : this(workbook, null, null)
        { }


        private XSSFFormulaEvaluator(XSSFWorkbook workbook, IStabilityClassifier stabilityClassifier, UDFFinder udfFinder)
            : this(workbook, new WorkbookEvaluator(XSSFEvaluationWorkbook.Create(workbook), stabilityClassifier, udfFinder))
        {
        }

        protected XSSFFormulaEvaluator(XSSFWorkbook workbook, WorkbookEvaluator bookEvaluator)
            : base(bookEvaluator)
        {
            _book = workbook;
        }

        /**
         * @param stabilityClassifier used to optimise caching performance. Pass <code>null</code>
         * for the (conservative) assumption that any cell may have its defInition Changed After
         * Evaluation begins.
         * @param udfFinder pass <code>null</code> for default (AnalysisToolPak only)
         */
        public static XSSFFormulaEvaluator Create(XSSFWorkbook workbook, IStabilityClassifier stabilityClassifier, UDFFinder udfFinder)
        {
            return new XSSFFormulaEvaluator(workbook, stabilityClassifier, udfFinder);
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
        public static void EvaluateAllFormulaCells(XSSFWorkbook wb)
        {
            BaseFormulaEvaluator.EvaluateAllFormulaCells(wb);
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
        public override void EvaluateAll()
        {
            EvaluateAllFormulaCells(_book, this);
        }

        /**
	     * Turns a XSSFCell into a XSSFEvaluationCell
	     */
        protected override IEvaluationCell ToEvaluationCell(ICell cell)
        {
            if (!(cell is XSSFCell)){
                throw new ArgumentException("Unexpected type of cell: " + cell.GetType().Name + "." +
                        " Only XSSFCells can be evaluated.");
            }

            return new XSSFEvaluationCell((XSSFCell)cell);
        }
    }
}