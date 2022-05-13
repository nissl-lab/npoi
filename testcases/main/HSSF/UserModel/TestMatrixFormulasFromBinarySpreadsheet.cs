using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.HSSF.UserModel
{
    [TestFixture]
    public class TestMatrixFormulasFromBinarySpreadsheet
    {
        private static HSSFWorkbook workbook;
        private static ISheet sheet;
        private static IFormulaEvaluator evaluator;

        private static class Navigator
        {
            /**
             * Name of the test spreadsheet (found in the standard test data folder)
             */
            public const String FILENAME = "MatrixFormulaEvalTestData.xls";
            /**
             * Row (zero-based) in the spreadsheet where operations start
             */
            public const int START_OPERATORS_ROW_INDEX = 1;
            /**
             * Column (zero-based) in the spreadsheet where operations start
             */
            public const int START_OPERATORS_COL_INDEX = 0;
            /**
             * Column (zero-based) in the spreadsheet where evaluations start
             */
            public const int START_RESULT_COL_INDEX = 7;
            /**
             * Column separation in the spreadsheet between evaluations and expected results
             */
            public const int COL_OFF_EXPECTED_RESULT = 3;
            /**
             * Row separation in the spreadsheet between operations
             */
            public const int ROW_OFF_NEXT_OP = 4;
            /**
             * Used to indicate when there are no more operations left
             */
            public const String END_OF_TESTS = "<END>";

        }


        public static List<Object[]> data()
        {
            // Function "Text" uses custom-formats which are locale specific
            // can't set the locale on a per-testrun execution, as some settings have been
            // already set, when we would try to change the locale by then

            workbook = HSSFTestDataSamples.OpenSampleWorkbook(Navigator.FILENAME);
            sheet = workbook.GetSheetAt(0);
            evaluator = new HSSFFormulaEvaluator(workbook);

            List<Object[]> data = new List<Object[]>();

            processFunctionGroup(data, Navigator.START_OPERATORS_ROW_INDEX, null);
            
            return data;
        }
    /**
 * @param startRowIndex row index in the spreadsheet where the first function/operator is found
 * @param testFocusFunctionName name of a single function/operator to test alone.
 * Typically pass <code>null</code> to test all functions
 */
    private static void processFunctionGroup(List<Object[]> data, int startRowIndex, String testFocusFunctionName)
        {
            for (int rowIndex = startRowIndex; true; rowIndex += Navigator.ROW_OFF_NEXT_OP)
            {
                IRow r = sheet.GetRow(rowIndex);
                String targetFunctionName = getTargetFunctionName(r);
                Assert.IsNotNull("Test spreadsheet cell empty on row ("
                        + (rowIndex) + "). Expected function name or '"
                        + Navigator.END_OF_TESTS + "'", targetFunctionName);
                if (targetFunctionName.Equals(Navigator.END_OF_TESTS))
                {
                    // found end of functions list
                    break;
                }
                if (testFocusFunctionName == null || targetFunctionName.Equals(testFocusFunctionName, StringComparison.InvariantCultureIgnoreCase))
                {
                    data.Add(new Object[] { targetFunctionName, rowIndex });
                }
            }
        }
        [TestCaseSource(nameof(data))]
        public void processFunctionRow(String targetFunctionName, int formulasRowIdx)
        {
            int endColNum = Navigator.START_RESULT_COL_INDEX + Navigator.COL_OFF_EXPECTED_RESULT;

            for (int rowNum = formulasRowIdx; rowNum < formulasRowIdx + Navigator.ROW_OFF_NEXT_OP - 1; rowNum++)
            {
                for (int colNum = Navigator.START_RESULT_COL_INDEX; colNum < endColNum; colNum++)
                {
                    IRow r = sheet.GetRow(rowNum);

                    /* mainly to escape row failures on MDETERM which only returns a scalar */
                    if (r == null)
                    {
                        continue;
                    }

                    ICell c = sheet.GetRow(rowNum).GetCell(colNum);

                    if (c == null || c.CellType != CellType.Formula)
                    {
                        continue;
                    }

                    CellValue actValue = evaluator.Evaluate(c);
                    ICell expValue = sheet.GetRow(rowNum).GetCell(colNum + Navigator.COL_OFF_EXPECTED_RESULT);

                    String msg = String.Format("Function '{0}': Formula: {1} @ {2}:{3}"
                            , targetFunctionName, c.CellFormula, rowNum, colNum);

                    Assert.IsNotNull(expValue, msg + " - Bad setup data expected value is null");
                    Assert.IsNotNull(actValue, msg + " - actual value was null");

                    CellType cellType = expValue.CellType;
                    switch (cellType)
                    {
                        case CellType.Blank:
                            Assert.AreEqual(CellType.Blank, actValue.CellType, msg);
                            break;
                        case CellType.Boolean:
                            Assert.AreEqual(CellType.Boolean, actValue.CellType, msg);
                            Assert.AreEqual(expValue.BooleanCellValue, actValue.BooleanValue, msg);
                            break;
                        case CellType.Error:
                            Assert.AreEqual(CellType.Error, actValue.CellType, msg);
                            Assert.AreEqual(ErrorEval.GetText(expValue.ErrorCellValue), ErrorEval.GetText(actValue.ErrorValue), msg);
                            break;
                        case CellType.Formula: // will never be used, since we will call method after formula evaluation
                            Assert.Fail("Cannot expect formula as result of formula evaluation: " + msg);
                            break;
                        case CellType.Numeric:
                            Assert.AreEqual(CellType.Numeric, actValue.CellType, msg);
                            Assert.AreEqual(expValue.NumericCellValue, actValue.NumberValue, msg);
                            break;
                        case CellType.String:
                            Assert.AreEqual(CellType.String, actValue.CellType, msg);
                            Assert.AreEqual(expValue.RichStringCellValue.String, actValue.StringValue, msg);
                            break;
                        default:
                            Assert.Fail("Unexpected cell type: " + cellType);
                            break;
                    }
                }
            }
        }
        /**
    * @return <code>null</code> if cell is missing, empty or Blank
    */
        private static String getTargetFunctionName(IRow r)
        {
            if (r == null)
            {
                Assert.Warn("Warning - given null row, can't figure out function name");
                return null;
            }
            ICell cell = r.GetCell(Navigator.START_OPERATORS_COL_INDEX);
            //Assert.Warn(Navigator.START_OPERATORS_COL_INDEX.ToString());
            if (cell == null)
            {
                Assert.Warn("Warning - Row " + r.RowNum + " has no cell " + Navigator.START_OPERATORS_COL_INDEX + ", can't figure out function name");
                return null;
            }
            if (cell.CellType == CellType.Blank)
            {
                return null;
            }
            if (cell.CellType == CellType.String)
            {
                return cell.RichStringCellValue.String;
            }
            throw new AssertionException("Bad cell type for 'function name' column: ("
                    + cell.CellType + ") row (" + (r.RowNum + 1) + ")");
        }
    }
}