/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/

using System;
using NPOI.SS.UserModel;
using NUnit.Framework;
using TestCases.SS.Formula.Functions;
using NPOI.OpenXml4Net.OPC;
using System.IO;
using TestCases.HSSF;
using NPOI.Util;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace TestCases.XSSF.UserModel
{


    /**
     * Performs much the same role as {@link TestFormulasFromSpreadsheet},
     *  except for a XSSF spreadsheet, not a HSSF one.
     * This allows us to check that all our Formula Evaluation code
     *  is able to work for XSSF, as well as for HSSF.
     * 
     * Periodically, you should open FormulaEvalTestData.xls in
     *  Excel 2007, and re-save it as FormulaEvalTestData_Copy.xlsx
     *  
     */
    [TestFixture]
    public class TestFormulaEvaluatorOnXSSF
    {

        private class Result
        {
            public const int SOME_EVALUATIONS_FAILED = -1;
            public const int ALL_EVALUATIONS_SUCCEEDED = +1;
            public const int NO_EVALUATIONS_FOUND = 0;
        }

        /** 
         * This class defines constants for navigating around the Test data spreadsheet used for these Tests.
         */
        private static class SS
        {

            /**
             * Name of the Test spreadsheet (found in the standard Test data folder)
             */
            public static String FILENAME = "FormulaEvalTestData_Copy.xlsx";
            /**
             * Row (zero-based) in the Test spreadsheet where the operator examples start.
             */
            public static int START_OPERATORS_ROW_INDEX = 22; // Row '23'
            /**
             * Row (zero-based) in the Test spreadsheet where the function examples start.
             */
            public static int START_FUNCTIONS_ROW_INDEX = 95; // Row '96'
            /**
             * Index of the column that Contains the function names
             */
            public static int COLUMN_INDEX_FUNCTION_NAME = 1; // Column 'B'

            /**
             * Used to indicate when there are no more functions left
             */
            public static String FUNCTION_NAMES_END_SENTINEL = "<END-OF-FUNCTIONS>";

            /**
             * Index of the column where the Test values start (for each function)
             */
            public static short COLUMN_INDEX_FIRST_TEST_VALUE = 3; // Column 'D'

            /**
             * Each function takes 4 rows in the Test spreadsheet
             */
            public static int NUMBER_OF_ROWS_PER_FUNCTION = 4;
        }

        private XSSFWorkbook workbook;
        private ISheet sheet;
        // Note - multiple failures are aggregated before ending.  
        // If one or more functions fail, a single AssertionFailedError is thrown at the end
        private int _functionFailureCount;
        private int _functionSuccessCount;
        private int _EvaluationFailureCount;
        private int _EvaluationSuccessCount;

        private static ICell GetExpectedValueCell(IRow row, short columnIndex)
        {
            if (row == null)
            {
                return null;
            }
            return row.GetCell(columnIndex);
        }


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
                    Assert.AreEqual(CellType.Boolean, actual.CellType, msg);
                    Assert.AreEqual(expected.BooleanCellValue, actual.BooleanValue, msg);
                    break;
                case CellType.Error:
                    Assert.AreEqual(CellType.Error, actual.CellType, msg);
                    //if (false)
                    //{ // TODO: fix ~45 functions which are currently returning incorrect error values
                    //	Assert.AreEqual(expected.ErrorCellValue, actual.ErrorValue, msg);
                    //}
                    break;
                case CellType.Formula: // will never be used, since we will call method After formula Evaluation
                    throw new AssertionException("Cannot expect formula as result of formula Evaluation: " + msg);
                case CellType.Numeric:
                    Assert.AreEqual(CellType.Numeric, actual.CellType, msg);
                    AbstractNumericTestCase.AssertEquals(msg, expected.NumericCellValue, actual.NumberValue, TestMathX.POS_ZERO, TestMathX.DIFF_TOLERANCE_FACTOR);

                    //				double delta = Math.abs(expected.NumericCellValue-actual.NumberValue);
                    //				double pctExpected = Math.abs(0.00001*expected.NumericCellValue);
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
                sheet = workbook.GetSheetAt(0);
            }
            _functionFailureCount = 0;
            _functionSuccessCount = 0;
            _EvaluationFailureCount = 0;
            _EvaluationSuccessCount = 0;
        }

        /**
         * Checks that we can actually open the file
         */
        [Test]
        public void TestOpen()
        {
            Assert.IsNotNull(workbook);
        }

        /**
         * Disabled for now, as many things seem to break
         *  for XSSF, which is a shame
         */
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestFunctionsFromTestSpreadsheet()
        {
            NumberToTextConverter.RawDoubleBitsToText(BitConverter.DoubleToInt64Bits(1.0));
            ProcessFunctionGroup(SS.START_OPERATORS_ROW_INDEX, null);
            ProcessFunctionGroup(SS.START_FUNCTIONS_ROW_INDEX, null);
            // example for debugging individual functions/operators:
            //		ProcessFunctionGroup(SS.START_OPERATORS_ROW_INDEX, "ConcatEval");
            //		ProcessFunctionGroup(SS.START_FUNCTIONS_ROW_INDEX, "AVERAGE");

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
            //if (false)
            //{ // normally no output for successful Tests
            //	Console.WriteLine(this.GetType().Name + ": " + successMsg);
            //}
        }

        /**
         * @param startRowIndex row index in the spreadsheet where the first function/operator is found 
         * @param TestFocusFunctionName name of a single function/operator to Test alone. 
         * Typically pass <code>null</code> to Test all functions
         */
        private void ProcessFunctionGroup(int startRowIndex, String testFocusFunctionName)
        {

            IFormulaEvaluator evaluator = new XSSFFormulaEvaluator(workbook);

            int rowIndex = startRowIndex;
            while (true)
            {
                IRow r = sheet.GetRow(rowIndex);
                String targetFunctionName = GetTargetFunctionName(r);
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
                if (testFocusFunctionName == null || targetFunctionName.Equals(testFocusFunctionName, StringComparison.OrdinalIgnoreCase))
                {

                    // expected results are on the row below
                    IRow expectedValuesRow = sheet.GetRow(rowIndex + 1);
                    if (expectedValuesRow == null)
                    {
                        int missingRowNum = rowIndex + 2; //+1 for 1-based, +1 for next row
                        throw new AssertionException("Missing expected values row for function '"
                                + targetFunctionName + " (row " + missingRowNum + ")");
                    }
                    switch (ProcessFunctionRow(evaluator, targetFunctionName, r, expectedValuesRow))
                    {
                        case Result.ALL_EVALUATIONS_SUCCEEDED: _functionSuccessCount++; break;
                        case Result.SOME_EVALUATIONS_FAILED: _functionFailureCount++; break;
                        case Result.NO_EVALUATIONS_FOUND: // do nothing
                            break;
                        default:
                            throw new RuntimeException("unexpected result");

                    }
                }
                rowIndex += SS.NUMBER_OF_ROWS_PER_FUNCTION;
            }
        }

        /**
         * 
         * @return a constant from the local Result class denoting whether there were any Evaluation
         * cases, and whether they all succeeded.
         */
        private int ProcessFunctionRow(IFormulaEvaluator Evaluator, String targetFunctionName,
                IRow formulasRow, IRow expectedValuesRow)
        {

            int result = Result.NO_EVALUATIONS_FOUND; // so far
            short endcolnum = formulasRow.LastCellNum;

            // iterate across the row for all the Evaluation cases
            for (short colnum = SS.COLUMN_INDEX_FIRST_TEST_VALUE; colnum < endcolnum; colnum++)
            {
                ICell c = formulasRow.GetCell(colnum);
                if (c == null || c.CellType != CellType.Formula)
                {
                    continue;
                }
                if (IsIgnoredFormulaTestCase(c.CellFormula))
                {
                    continue;
                }

                CellValue actualValue;
                try
                {
                    actualValue = Evaluator.Evaluate(c);
                }
                catch (RuntimeException e)
                {
                    _EvaluationFailureCount++;
                    PrintshortStackTrace(System.Console.Error, e);
                    result = Result.SOME_EVALUATIONS_FAILED;
                    continue;
                }

                ICell expectedValueCell = GetExpectedValueCell(expectedValuesRow, colnum);
                try
                {
                    ConfirmExpectedResult("Function '" + targetFunctionName + "': Formula: " + c.CellFormula + " @ " + formulasRow.RowNum + ":" + colnum,
                            expectedValueCell, actualValue);
                    _EvaluationSuccessCount++;
                    if (result != Result.SOME_EVALUATIONS_FAILED)
                    {
                        result = Result.ALL_EVALUATIONS_SUCCEEDED;
                    }
                }
                catch (AssertionException e)
                {
                    _EvaluationFailureCount++;
                    PrintshortStackTrace(System.Console.Error, e);
                    result = Result.SOME_EVALUATIONS_FAILED;
                }
            }
            return result;
        }

        /*
         * TODO - these are all formulas which currently (Apr-2008) break on ooxml 
         */
        private static bool IsIgnoredFormulaTestCase(String cellFormula)
        {
            if ("COLUMN(1:2)".Equals(cellFormula) || "ROW(2:3)".Equals(cellFormula))
            {
                // full row ranges are not Parsed properly yet.
                // These cases currently work in svn tRunk because of another bug which causes the 
                // formula to Get rendered as COLUMN($A$1:$IV$2) or ROW($A$2:$IV$3) 
                return true;
            }
            if ("ISREF(currentcell())".Equals(cellFormula))
            {
                // currently throws NPE because unknown function "currentcell" causes name lookup 
                // Name lookup requires some equivalent object of the Workbook within xSSFWorkbook.
                return true;
            }
            return false;
        }


        /**
         * Useful to keep output concise when expecting many failures to be reported by this Test case
         */
        private static void PrintshortStackTrace(TextWriter ps, Exception e)
        {
            //StackTraceElement[] stes = e.GetStackTrace();

            //int startIx = 0;
            //// Skip any top frames inside junit.framework.Assert
            //while(startIx<stes.Length) {
            //    if(!stes[startIx].GetClassName().Equals(Assert.class.GetName())) {
            //        break;
            //    }
            //    startIx++;
            //}
            //// Skip bottom frames (part of junit framework)
            //int endIx = startIx+1;
            //while(endIx < stes.Length) {
            //    if(stes[endIx].GetClassName().Equals(TestCase.class.GetName())) {
            //        break;
            //    }
            //    endIx++;
            //}
            //if(startIx >= endIx) {
            //    // something went wrong. just print the whole stack trace
            //    e.printStackTrace(ps);
            //}
            //endIx -= 4; // Skip 4 frames of reflection invocation
            //ps.println(e.ToString());
            //for(int i=startIx; i<endIx; i++) {
            //    ps.println("\tat " + stes[i].ToString());
            //}
        }

        /**
         * @return <code>null</code> if cell is missing, empty or blank
         */
        private static String GetTargetFunctionName(IRow r)
        {
            if (r == null)
            {
                System.Console.WriteLine("Warning - given null row, can't figure out function name");
                return null;
            }
            ICell cell = r.GetCell(SS.COLUMN_INDEX_FUNCTION_NAME);
            if (cell == null)
            {
                System.Console.WriteLine("Warning - Row " + r.RowNum + " has no cell " + SS.COLUMN_INDEX_FUNCTION_NAME + ", can't figure out function name");
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