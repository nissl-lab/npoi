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

namespace TestCases.XSSF.UserModel
{
    using System;
    using System.IO;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.UserModel;
    using NPOI.Util;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using TestCases.HSSF;
    using TestCases.SS.Formula.Eval;
    using TestCases.SS.Formula.Functions;


    /**
     * Tests formulas for multi sheet reference (i.e. SUM(Sheet1:Sheet5!A1))
     */
    [TestFixture]
    public class TestMultiSheetFormulaEvaluatorOnXSSF
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(TestFormulasFromSpreadsheet));

        private static class Result
        {
            public const int SOME_EVALUATIONS_FAILED = -1;
            public const int ALL_EVALUATIONS_SUCCEEDED = +1;
            public const int NO_EVALUATIONS_FOUND = 0;
        }

        /**
         * This class defines constants for navigating around the test data spreadsheet used for these tests.
         */
        private static class SS
        {

            /**
             * Name of the test spreadsheet (found in the standard test data folder)
             */
            public static String FILENAME = "FormulaSheetRange.xlsx";
            /**
             * Row (zero-based) in the test spreadsheet where the function examples start.
             */
            public static int START_FUNCTIONS_ROW_INDEX = 10; // Row '11'
            /**
             * Index of the column that Contains the function names
             */
            public static int COLUMN_INDEX_FUNCTION_NAME = 0; // Column 'A'
            /**
             * Index of the column that Contains the test names
             */
            public static int COLUMN_INDEX_TEST_NAME = 1; // Column 'B'
            /**
             * Used to indicate when there are no more functions left
             */
            public static String FUNCTION_NAMES_END_SENTINEL = "<END>";

            /**
             * Index of the column where the test expected value is present
             */
            public static short COLUMN_INDEX_EXPECTED_VALUE = 2; // Column 'C'
            /**
             * Index of the column where the test actual value is present
             */
            public static short COLUMN_INDEX_ACTUAL_VALUE = 3; // Column 'D'
            /**
             * Test sheet name (sheet with all test formulae)
             */
            public static String TEST_SHEET_NAME = "test";
        }

        private XSSFWorkbook workbook;
        private ISheet sheet;
        // Note - multiple failures are aggregated before ending.
        // If one or more functions fail, a single AssertFailedException is thrown at the end
        private int _functionFailureCount;
        private int _functionSuccessCount;
        private int _EvaluationFailureCount;
        private int _EvaluationSuccessCount;

        private static void ConfirmExpectedResult(String msg, ICell expected, CellValue actual)
        {
            if (expected == null)
            {
                throw new AssertionException(msg + " - Bad Setup data expected value is null");
            }
            if (actual == null)
            {
                throw new AssertionException(msg + " - actual value was null");
            }

            switch (expected.CellType)
            {
                case CellType.Blank:
                    Assert.AreEqual(CellType.Blank, actual.CellType, msg);
                    break;
                case CellType.Boolean:
                    Assert.AreEqual(CellType.Boolean, actual.CellType);
                    Assert.AreEqual(expected.BooleanCellValue, actual.BooleanValue, msg);
                    break;
                case CellType.Error:
                    Assert.AreEqual(CellType.Error, actual.CellType, msg);
                    if (false)
                    { // TODO: fix ~45 functions which are currently returning incorrect error values
                        // Assert.AreEqual(expected.ErrorCellValue, actual.ErrorValue, msg);
                    }
                    break;
                case CellType.Formula: // will never be used, since we will call method After formula Evaluation
                    throw new AssertionException("Cannot expect formula as result of formula Evaluation: " + msg);
                case CellType.Numeric:
                    Assert.AreEqual(CellType.Numeric, actual.CellType, msg);
                    TestMathX.AssertEquals(msg, expected.NumericCellValue, actual.NumberValue, TestMathX.POS_ZERO, TestMathX.DIFF_TOLERANCE_FACTOR);
                    //				double delta = Math.Abs(expected.NumericCellValue-actual.NumberValue);
                    //				double pctExpected = Math.Abs(0.00001*expected.NumericCellValue);
                    //				Assert.IsTrue(msg, delta <= pctExpected);
                    break;
                case CellType.String:
                    Assert.AreEqual(CellType.String, actual.CellType, msg);
                    Assert.AreEqual(expected.RichStringCellValue.String, actual.StringValue, msg);
                    break;
            }
        }

        [SetUp]
        public void SetUp()
        {
            if (workbook == null)
            {
                Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(SS.FILENAME);
                OPCPackage pkg = OPCPackage.Open(is1);
                workbook = new XSSFWorkbook(pkg);
                sheet = workbook.GetSheet(SS.TEST_SHEET_NAME);
            }
            _functionFailureCount = 0;
            _functionSuccessCount = 0;
            _EvaluationFailureCount = 0;
            _EvaluationSuccessCount = 0;
        }

        [Test]
        public void TestFunctionsFromTestSpreadsheet()
        {

            ProcessFunctionGroup(SS.START_FUNCTIONS_ROW_INDEX, null);

            // confirm results
            String successMsg = "There were "
                    + _EvaluationSuccessCount + " successful Evaluation(s) and "
                    + _functionSuccessCount + " function(s) without error";
            if (_functionFailureCount > 0)
            {
                String msg = _functionFailureCount + " function(s) failed in "
                + _EvaluationFailureCount + " Evaluation(s).  " + successMsg;
                throw new AssertionException(msg);
            }
            logger.Log(POILogger.INFO, GetType().Name + ": " + successMsg);
        }

        /**
         * @param startRowIndex row index in the spreadsheet where the first function/operator is found
         * @param testFocusFunctionName name of a single function/operator to test alone.
         * Typically pass <code>null</code> to test all functions
         */
        private void ProcessFunctionGroup(int startRowIndex, String testFocusFunctionName)
        {
            IFormulaEvaluator evaluator = new XSSFFormulaEvaluator(workbook);

            int rowIndex = startRowIndex;
            while (true)
            {
                IRow r = sheet.GetRow(rowIndex);

                // only Evaluate non empty row
                if (r != null)
                {
                    String targetFunctionName = GetTargetFunctionName(r);
                    String targetTestName = GetTargetTestName(r);
                    if (targetFunctionName == null)
                    {
                        throw new AssertionException("Test spreadsheet cell empty on row ("
                                + (rowIndex + 1) + "). Expected function name or '"
                                + SS.FUNCTION_NAMES_END_SENTINEL + "'");
                    }
                    if (targetFunctionName.Equals(SS.FUNCTION_NAMES_END_SENTINEL))
                    {
                        // found end of functions list
                        break;
                    }
                    if (testFocusFunctionName == null || targetFunctionName.Equals(testFocusFunctionName, StringComparison.CurrentCultureIgnoreCase))
                    {

                        // expected results are on the row below
                        ICell expectedValueCell = r.GetCell(SS.COLUMN_INDEX_EXPECTED_VALUE);
                        if (expectedValueCell == null)
                        {
                            int missingRowNum = rowIndex + 1;
                            throw new AssertionException("Missing expected values cell for function '"
                                    + targetFunctionName + ", test" + targetTestName + " (row " +
                                    missingRowNum + ")");
                        }

                        switch (ProcessFunctionRow(evaluator, targetFunctionName, targetTestName, r, expectedValueCell))
                        {
                            case Result.ALL_EVALUATIONS_SUCCEEDED:
                                _functionSuccessCount++;
                                break;
                            case Result.SOME_EVALUATIONS_FAILED:
                                _functionFailureCount++;
                                break;
                            default:
                                throw new Exception("unexpected result");
                            case Result.NO_EVALUATIONS_FOUND: // do nothing
                                break;
                        }
                    }
                }
                rowIndex++;
            }
        }

        /**
         *
         * @return a constant from the local Result class denoting whether there were any Evaluation
         * cases, and whether they all succeeded.
         */
        private int ProcessFunctionRow(IFormulaEvaluator Evaluator, String targetFunctionName,
                String targetTestName, IRow formulasRow, ICell expectedValueCell)
        {

            int result = Result.NO_EVALUATIONS_FOUND; // so far

            ICell c = formulasRow.GetCell(SS.COLUMN_INDEX_ACTUAL_VALUE);
            if (c == null || c.CellType != CellType.Formula)
            {
                return result;
            }

            CellValue actualValue = Evaluator.Evaluate(c);

            try
            {
                ConfirmExpectedResult("Function '" + targetFunctionName + "': Test: '" + targetTestName + "' Formula: " + c.CellFormula
                + " @ " + formulasRow.RowNum + ":" + SS.COLUMN_INDEX_ACTUAL_VALUE,
                        expectedValueCell, actualValue);
                _EvaluationSuccessCount++;
                if (result != Result.SOME_EVALUATIONS_FAILED)
                {
                    result = Result.ALL_EVALUATIONS_SUCCEEDED;
                }
            }
            catch (Exception e)
            {
                _EvaluationFailureCount++;
                //printshortStackTrace(System.err, e);
                Console.WriteLine(e.StackTrace);
                result = Result.SOME_EVALUATIONS_FAILED;
            }

            return result;
        }

        /**
         * Useful to keep output concise when expecting many failures to be reported by this test case
         */
        private static void printshortStackTrace(TextWriter ps, AssertionException e)
        {
            //StackTraceElement[] stes = e.StackTrace;

            //int startIx = 0;
            //// skip any top frames inside junit.framework.Assert
            //while (startIx < stes.Length)
            //{
            //    if (!stes[startIx].ClassName.Equals(typeof(Assert).Name))
            //    {
            //        break;
            //    }
            //    startIx++;
            //}
            //// skip bottom frames (part of junit framework)
            //int endIx = startIx + 1;
            //while (endIx < stes.Length)
            //{
            //    if (stes[endIx].ClassName.Equals(typeof(TestCase).Name))
            //    {
            //        break;
            //    }
            //    endIx++;
            //}
            //if (startIx >= endIx)
            //{
            //    // something went wrong. just print the whole stack trace
            //    e.PrintStackTrace(ps);
            //}
            //endIx -= 4; // skip 4 frames of reflection invocation
            //ps.Println(e.ToString());
            //for (int i = startIx; i < endIx; i++)
            //{
            //    ps.Println("\tat " + stes[i].ToString());
            //}
        }

        /**
         * @return <code>null</code> if cell is missing, empty or blank
         */
        private static String GetTargetFunctionName(IRow r)
        {
            if (r == null)
            {
                Console.WriteLine("Warning - given null row, can't figure out function name");
                return null;
            }
            ICell cell = r.GetCell(SS.COLUMN_INDEX_FUNCTION_NAME);
            if (cell == null)
            {
                Console.WriteLine("Warning - Row " + r.RowNum + " has no cell " + SS.COLUMN_INDEX_FUNCTION_NAME + ", can't figure out function name");
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
        /**
         * @return <code>null</code> if cell is missing, empty or blank
         */
        private static String GetTargetTestName(IRow r)
        {
            if (r == null)
            {
                Console.WriteLine("Warning - given null row, can't figure out test name");
                return null;
            }
            ICell cell = r.GetCell(SS.COLUMN_INDEX_TEST_NAME);
            if (cell == null)
            {
                Console.WriteLine("Warning - Row " + r.RowNum + " has no cell " + SS.COLUMN_INDEX_TEST_NAME + ", can't figure out test name");
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

            throw new AssertionException("Bad cell type for 'test name' column: ("
                    + cell.CellType + ") row (" + (r.RowNum + 1) + ")");
        }

    }

}