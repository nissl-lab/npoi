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

using System.Collections.ObjectModel;

namespace TestCases.SS.Formula.Eval
{

    using System;
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;
    using TestCases.SS.Formula.Functions;
    using System.Diagnostics;

    /**
     * Tests formulas and operators as loaded from a Test data spreadsheet.<p/>
     * This class does not Test implementors of <c>Function</c> and <c>OperationEval</c> in
     * isolation.  Much of the Evaluation engine (i.e. <c>HSSFFormulaEvaluator</c>, ...) Gets
     * exercised as well.  Tests for bug fixes and specific/tricky behaviour can be found in the
     * corresponding Test class (<c>TestXxxx</c>) of the target (<c>Xxxx</c>) implementor,
     * where execution can be observed more easily.
     *
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     */
    [TestFixture]
    public class TestFormulasFromSpreadsheet
    {

        private static class Result
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
            public static String FILENAME = "FormulaEvalTestData.xls";
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

        private HSSFWorkbook workbook;
        private ISheet sheet;
        // Note - multiple failures are aggregated before ending.
        // If one or more functions fail, a single AssertionException is thrown at the end
        private int _functionFailureCount;
        private int _functionSuccessCount;
        private int _EvaluationFailureCount;
        private int _EvaluationSuccessCount;

        private static ICell GetExpectedValueCell(IRow row, int columnIndex)
        {
            if (row == null)
            {
                return null;
            }
            return row.GetCell(columnIndex);
        }


        private static void ConfirmExpectedResult(String msg, ICell expected, CellValue actual)
        {
            Assert.IsNotNull(expected, msg + " - Bad setup data expected value is null");
            Assert.IsNotNull(actual, msg + " - actual value was null");
            
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
                    Assert.AreEqual(ErrorEval.GetText(expected.ErrorCellValue), ErrorEval.GetText(actual.ErrorValue), msg);
                    break;
                case CellType.Formula: // will never be used, since we will call method After formula Evaluation
                    Assert.Fail("Cannot expect formula as result of formula Evaluation: " + msg);
                    break;
                case CellType.Numeric:
                    Assert.AreEqual(CellType.Numeric, actual.CellType, msg);
                    AbstractNumericTestCase.AssertEquals(msg, expected.NumericCellValue, actual.NumberValue,
                        AbstractNumericTestCase.POS_ZERO, AbstractNumericTestCase.DIFF_TOLERANCE_FACTOR);
                    break;
                case CellType.String:
                    Assert.AreEqual(CellType.String, actual.CellType, msg);
                    Assert.AreEqual(expected.RichStringCellValue.String, actual.StringValue, msg);
                    break;
            }
        }

        [SetUp]
        protected void SetUp()
        {
            if (workbook == null)
            {
                workbook = HSSFTestDataSamples.OpenSampleWorkbook(SS.FILENAME);
                sheet = workbook.GetSheetAt(0);
            }
            _functionFailureCount = 0;
            _functionSuccessCount = 0;
            _EvaluationFailureCount = 0;
            _EvaluationSuccessCount = 0;
        }
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestFunctionsFromTestSpreadsheet()
        {

            ProcessFunctionGroup(SS.START_OPERATORS_ROW_INDEX, null);
            ProcessFunctionGroup(SS.START_FUNCTIONS_ROW_INDEX, null);
            // example for debugging individual functions/operators:
            //		ProcessFunctionGroup(SS.START_OPERATORS_ROW_INDEX, "ConcatEval");
            //		ProcessFunctionGroup(SS.START_FUNCTIONS_ROW_INDEX, "AVERAGE");

            // confirm results
            String successMsg = "There were "
                    + _EvaluationSuccessCount + " successful Evaluation(s) and "
                    + _functionSuccessCount + " function(s) without error";
            
            String msg = _functionFailureCount + " function(s) failed in "
            + _EvaluationFailureCount + " Evaluation(s).  " + successMsg;
            Assert.AreEqual(_functionFailureCount, 0, msg);


            Debug.WriteLine(this.GetType().Name + ": " + successMsg);


        }

        /**
         * @param startRowIndex row index in the spreadsheet where the first function/operator is found
         * @param TestFocusFunctionName name of a single function/operator to Test alone.
         * Typically pass <code>null</code> to Test all functions
         */
        private void ProcessFunctionGroup(int startRowIndex, String testFocusFunctionName)
        {
            HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(workbook);
            ReadOnlyCollection<String> funcs = FunctionEval.GetSupportedFunctionNames();
            int rowIndex = startRowIndex;
            while (true)
            {
                IRow r = sheet.GetRow(rowIndex);
                String targetFunctionName = GetTargetFunctionName(r);
                Assert.IsNotNull(targetFunctionName, "Test spreadsheet cell empty on row ("
                            + (rowIndex + 1) + "). Expected function name or '"
                            + SS.FUNCTION_NAMES_END_SENTINEL + "'");
                
                if (targetFunctionName.Equals(SS.FUNCTION_NAMES_END_SENTINEL))
                {
                    // found end of functions list
                    break;
                }
                if (testFocusFunctionName == null || targetFunctionName.Equals(testFocusFunctionName, StringComparison.CurrentCultureIgnoreCase))
                {
                    // expected results are on the row below
                    IRow expectedValuesRow = sheet.GetRow(rowIndex + 1);
                    
                    int missingRowNum = rowIndex + 2; //+1 for 1-based, +1 for next row
                    Assert.IsNotNull(expectedValuesRow, "Missing expected values row for function '"
                            + targetFunctionName + " (row " + missingRowNum + ")");
                    
                    switch (ProcessFunctionRow(evaluator, targetFunctionName, r, expectedValuesRow))
                    {
                        case Result.ALL_EVALUATIONS_SUCCEEDED: _functionSuccessCount++; break;
                        case Result.SOME_EVALUATIONS_FAILED: _functionFailureCount++; break;
                        default:
                            throw new SystemException("unexpected result");
                        case Result.NO_EVALUATIONS_FOUND: // do nothing
                            String uname = targetFunctionName.ToUpper();
                            if (startRowIndex >= SS.START_FUNCTIONS_ROW_INDEX &&
                                    funcs.Contains(uname))
                            {
                                Debug.WriteLine(uname + ": function is supported but missing test data", "");
                            }
                            break;
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
        private int ProcessFunctionRow(HSSFFormulaEvaluator Evaluator, String targetFunctionName,
                IRow formulasRow, IRow expectedValuesRow)
        {

            int result = Result.NO_EVALUATIONS_FOUND; // so far
            short endcolnum = (short)formulasRow.LastCellNum;

            // iterate across the row for all the Evaluation cases
            for (int colnum = SS.COLUMN_INDEX_FIRST_TEST_VALUE; colnum < endcolnum; colnum++)
            {
                ICell c = formulasRow.GetCell(colnum);
                if (c == null || c.CellType != CellType.Formula)
                {
                    continue;
                }

                CellValue actualValue = Evaluator.Evaluate(c);

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
                    printshortStackTrace(System.Console.Error, e);
                    result = Result.SOME_EVALUATIONS_FAILED;
                }
            }
            return result;
        }

        /**
         * Useful to keep output concise when expecting many failures to be reported by this Test case
         */
        private static void printshortStackTrace(TextWriter ps, AssertionException e)
        {
            ps.WriteLine(e.Message);
            ps.WriteLine(e.StackTrace);
            /*StackTraceElement[] stes = e.GetStackTrace();

            int startIx = 0;
            // skip any top frames inside junit.framework.Assert
            while(startIx<stes.Length) {
                if(!stes[startIx].GetClassName().Equals(typeof(Assert).Name)) {
                    break;
                }
                startIx++;
            }
            // skip bottom frames (part of junit framework)
            int endIx = startIx+1;
            while(endIx < stes.Length) {
                if (stes[endIx].GetClassName().Equals(typeof(Assert).Name))
                {
                    break;
                }
                endIx++;
            }
            if(startIx >= endIx) {
                // something went wrong. just print the whole stack trace
                e.printStackTrace(ps);
            }
            endIx -= 4; // skip 4 frames of reflection invocation
            ps.println(e.ToString());
            for(int i=startIx; i<endIx; i++) {
                ps.println("\tat " + stes[i].ToString());
            }*/
        }

        /**
         * @return <code>null</code> if cell is missing, empty or blank
         */
        private static String GetTargetFunctionName(IRow r)
        {
            if (r == null)
            {
                System.Console.Error.WriteLine("Warning - given null row, can't figure out function name");
                return null;
            }
            ICell cell = r.GetCell(SS.COLUMN_INDEX_FUNCTION_NAME);
            if (cell == null)
            {
                System.Console.Error.WriteLine("Warning - Row " + r.RowNum + " has no cell " + SS.COLUMN_INDEX_FUNCTION_NAME + ", can't figure out function name");
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