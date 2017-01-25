using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.Udf;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.XSSF.Streaming
{
    public class SXSSFFormulaEvaluator : XSSFFormulaEvaluator
    {
        //TODO: Readonly?
        private static POILogger logger = POILogFactory.GetLogger(typeof(SXSSFFormulaEvaluator));

        private IWorkbook wb;

        public SXSSFFormulaEvaluator(SXSSFWorkbook workbook) : this(workbook, null, null)
        {
        }

        private SXSSFFormulaEvaluator(SXSSFWorkbook workbook, IStabilityClassifier stabilityClassifier, UDFFinder udfFinder) : this(workbook, new WorkbookEvaluator(SXSSFEvaluationWorkbook.Create(workbook), stabilityClassifier, udfFinder))
        {

        }

        //TODO: implemented new constructor
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
                if (((SXSSFSheet)sheet).AreAllRowsFlushed())
                {
                    throw new SheetsFlushedException();
                }
            }

            // Process the sheets as best we can
            foreach (ISheet sheet in wb)
            {

                // Check if any rows have already been flushed out
                int lastFlushedRowNum = ((SXSSFSheet)sheet).GetLastFlushedRowNum();
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
