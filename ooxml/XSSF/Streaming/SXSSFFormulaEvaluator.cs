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
using System;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Udf;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFFormulaEvaluator : XSSFFormulaEvaluator
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(SXSSFFormulaEvaluator));

        private IWorkbook wb;

        public SXSSFFormulaEvaluator(SXSSFWorkbook workbook) : this(workbook, null, null)
        {
        }

        private SXSSFFormulaEvaluator(SXSSFWorkbook workbook, IStabilityClassifier stabilityClassifier, UDFFinder udfFinder) : this(workbook, new WorkbookEvaluator(SXSSFEvaluationWorkbook.Create(workbook), stabilityClassifier, udfFinder))
        {

        }

        private SXSSFFormulaEvaluator(SXSSFWorkbook workbook, WorkbookEvaluator bookEvaluator) : base(bookEvaluator)
        {
            this.wb = workbook;
        }

        public static SXSSFFormulaEvaluator create(SXSSFWorkbook workbook, IStabilityClassifier stabilityClassifier, UDFFinder udfFinder)
        {
            return new SXSSFFormulaEvaluator(workbook, stabilityClassifier, udfFinder);
        }

        protected IEvaluationCell toEvaluationCell(ICell cell)
        {
            if (!(cell is SXSSFCell))
            {
                throw new ArgumentException("Unexpected type of cell: " + cell.GetType() + "." +
                        " Only SXSSFCells can be evaluated.");
            }

            return new SXSSFEvaluationCell((SXSSFCell)cell);
        }

        public SXSSFCell EvaluateInCell(ICell cell)
        {
            return (SXSSFCell)base.EvaluateInCell(cell);
        }

        public static void EvaluateAllFormulaCells(SXSSFWorkbook wb, bool skipOutOfWindow)
        {
            SXSSFFormulaEvaluator eval = new SXSSFFormulaEvaluator(wb);

            // Check they're all available
            foreach (ISheet sheet in wb)
            {
                if (((SXSSFSheet)sheet).allFlushed)
                {
                    throw new SheetsFlushedException();
                }
            }

            // Process the sheets as best we can
            foreach (ISheet sheet in wb)
            {

                // Check if any rows have already been flushed out
                int lastFlushedRowNum = ((SXSSFSheet)sheet).lastFlushedRowNumber;
                if (lastFlushedRowNum > -1)
                {
                    if (!skipOutOfWindow) throw new RowFlushedException(0);
                    logger.Log(POILogger.INFO, "Rows up to " + lastFlushedRowNum + " have already been flushed, skipping");
                }

                // Evaluate what we have
                foreach (IRow r in sheet)
                {
                    foreach (ICell c in r)
                    {
                        if (c.CellType == CellType.Formula)
                        {
                            eval.EvaluateFormulaCell(c);
                        }
                    }
                }
            }
        }

        public void EvaluateAll()
        {
            // Have the evaluation done, with exceptions
            EvaluateAllFormulaCells((SXSSFWorkbook)wb, false);
        }

    }


    public class SheetsFlushedException : Exception
    {
        public SheetsFlushedException() : base("One or more sheets have been flushed, cannot evaluate all cells")
        {

        }
    }

    public class RowFlushedException : Exception
    {
        public RowFlushedException(int rowNum) : base("Row " + rowNum + " has been flushed, cannot evaluate all cells")
        {

        }
    }
}
