/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is1 distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/

namespace TestCases.HSSF.Record.Formula.Eval
{
    using System;
    using System.Collections;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.HSSF.Record.Formula.Eval;
    using TestCases.HSSF;
    using NPOI.HSSF.Record.Formula.Functions;
    using NPOI.HSSF.UserModel;
    using TestCases.HSSF.Record.Formula.Functions;

    using NPOI.SS.UserModel;

    /**
     * Tests formulas and operators as loaded from a test data spReadsheet.<p/>
     * This class does not test implementors of <tt>Function</tt> and <tt>OperationEval</tt> in
     * isolation.  Much of the evaluation engine (i.e. <tt>HSSFFormulaEvaluator</tt>, ...) Gets
     * exercised as well.  Tests for bug fixes and specific/tricky behaviour can be found in the
     * corresponding test class (<tt>TestXxxx</tt>) of the target (<tt>Xxxx</tt>) implementor, 
     * where execution can be observed more easily.
     * 
     * @author Amol S. Deshmukh &lt; amolweb at ya hoo dot com &gt;
     */
    public class TestFormulasFromSpReadsheet
    {

        private static class Result
        {
            public const int SOME_EVALUATIONS_FAILED = -1;
            public const int ALL_EVALUATIONS_SUCCEEDED = +1;
            public const int NO_EVALUATIONS_FOUND = 0;
        }

        /** 
         * This class defines constants for navigating around the test data spReadsheet used for these tests.
         */
        private static class SS
        {

            /**
             * Name of the test spReadsheet (found in the standard test data folder)
             */
            public static String FILENAME = "FormulaEvalTestData.xls";
            /**
             * NPOI.SS.UserModel.Row (zero-based) in the test spReadsheet where the operator examples start.
             */
            public static int START_OPERATORS_ROW_INDEX = 22; // NPOI.SS.UserModel.Row '23'
            /**
             * NPOI.SS.UserModel.Row (zero-based) in the test spReadsheet where the function examples start.
             */
            public static int START_FUNCTIONS_ROW_INDEX = 87; // NPOI.SS.UserModel.Row '88' 
            /** 
             * Index of the column that contains the function names
             */
            public static short COLUMN_INDEX_FUNCTION_NAME = 1; // Column 'B'

            /**
             * Used to indicate when there are no more functions left
             */
            public static String FUNCTION_NAMES_END_SENTINEL = "<END-OF-FUNCTIONS>";

            /**
             * Index of the column where the test values start (for each function)
             */
            public static short COLUMN_INDEX_FIRST_TEST_VALUE = 3; // Column 'D'

            /**
             * Each function takes 4 rows in the test spReadsheet 
             */
            public static int NUMBER_OF_ROWS_PER_FUNCTION = 4;
        }

        private HSSFWorkbook workbook;
        private NPOI.SS.UserModel.ISheet sheet;
        // Note - multiple failures are aggregated before ending.  
        // If one or more functions fail, a single AssertFailedException is1 thrown at the end
        private int _functionFailureCount;
        private int _functionSuccessCount;
        private int _evaluationFailureCount;
        private int _evaluationSuccessCount;

        private static ICell GetExpectedValueCell(NPOI.SS.UserModel.IRow row, short columnIndex)
        {
            if (row == null)
            {
                return null;
            }
            return row.GetCell(columnIndex);
        }


        private static void ConfirmExpectedResult(String msg, ICell expected, NPOI.SS.UserModel.CellValue actual)
        {
            if (expected == null)
            {
                throw new AssertFailedException(msg + " - Bad setup data expected value is null");
            }
            if (actual == null)
            {
                throw new AssertFailedException(msg + " - actual value was null");
            }

            switch (expected.CellType)
            {
                case NPOI.SS.UserModel.CellType.BLANK:
                    Assert.AreEqual(NPOI.SS.UserModel.CellType.BLANK, actual.CellType, msg);
                    break;
                case NPOI.SS.UserModel.CellType.BOOLEAN:
                    Assert.AreEqual(NPOI.SS.UserModel.CellType.BOOLEAN, actual.CellType, msg);
                    Assert.AreEqual(expected.BooleanCellValue, actual.BooleanValue, msg);
                    break;
                case NPOI.SS.UserModel.CellType.ERROR:
                    Assert.AreEqual(NPOI.SS.UserModel.CellType.ERROR, actual.CellType, msg);
                    Assert.AreEqual(ErrorEval.GetText(expected.ErrorCellValue), ErrorEval.GetText(actual.ErrorValue), msg);
                    break;
                case NPOI.SS.UserModel.CellType.FORMULA: // will never be used, since we will call method after formula evaluation
                    throw new AssertFailedException("Cannot expect formula as result of formula evaluation: " + msg);
                case NPOI.SS.UserModel.CellType.NUMERIC:
                    Assert.AreEqual(NPOI.SS.UserModel.CellType.NUMERIC, actual.CellType, msg);
                    //Assert.AreEqual(msg, expected.NumericCellValue, actual.NumberValue, TestMathX.POS_ZERO, TestMathX.DIFF_TOLERANCE_FACTOR);
                    Assert.AreEqual(expected.NumericCellValue, actual.NumberValue,msg);
                    break;
                case NPOI.SS.UserModel.CellType.STRING:
                    Assert.AreEqual(NPOI.SS.UserModel.CellType.STRING, actual.CellType, msg);
                    Assert.AreEqual(expected.RichStringCellValue.String, actual.StringValue, msg);
                    break;
            }
        }


        protected void SetUp()
        {
            if (workbook == null)
            {
                workbook = HSSFTestDataSamples.OpenSampleWorkbook(SS.FILENAME);
                sheet = workbook.GetSheetAt(0);
            }
            _functionFailureCount = 0;
            _functionSuccessCount = 0;
            _evaluationFailureCount = 0;
            _evaluationSuccessCount = 0;
        }

        public void TestFunctionsFromTestSpReadsheet() {
		
		processFunctionGroup(SS.START_OPERATORS_ROW_INDEX, null);
		processFunctionGroup(SS.START_FUNCTIONS_ROW_INDEX, null);
		// example for debugging individual functions/operators:
//		processFunctionGroup(SS.START_OPERATORS_ROW_INDEX, "ConcatEval");
//		processFunctionGroup(SS.START_FUNCTIONS_ROW_INDEX, "AVERAGE");
		
		// Confirm results
		String successMsg = "There were " 
				+ _evaluationSuccessCount + " successful evaluation(s) and "
				+ _functionSuccessCount + " function(s) without error";
		if(_functionFailureCount > 0) {
			String msg = _functionFailureCount + " function(s) failed in "
			+ _evaluationFailureCount + " evaluation(s).  " + successMsg;
			throw new AssertFailedException(msg);
		}
        //if(false) { // normally no output for successful tests
        //    Console.WriteLine(GetType().Name + ": " + successMsg);
        //}
	}

        /**
         * @param startRowIndex row index in the spReadsheet where the first function/operator is1 found 
         * @param testFocusFunctionName name of a single function/operator to test alone. 
         * Typically pass <c>null</c> to test all functions
         */
        private void processFunctionGroup(int startRowIndex, String testFocusFunctionName)
        {

            HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(sheet, workbook);

            int rowIndex = startRowIndex;
            while (true)
            {
                NPOI.SS.UserModel.IRow r = sheet.GetRow(rowIndex);
                String targetFunctionName = GetTargetFunctionName(r);
                if (targetFunctionName == null)
                {
                    throw new AssertFailedException("Test spReadsheet cell empty on row ("
                            + (rowIndex + 1) + "). Expected function name or '"
                            + SS.FUNCTION_NAMES_END_SENTINEL + "'");
                }
                if (targetFunctionName.Equals(SS.FUNCTION_NAMES_END_SENTINEL))
                {
                    // found end of functions list
                    break;
                }
                if (testFocusFunctionName == null || targetFunctionName.Equals(testFocusFunctionName, StringComparison.InvariantCultureIgnoreCase))
                {

                    // expected results are on the row below
                    NPOI.SS.UserModel.IRow expectedValuesRow = sheet.GetRow(rowIndex + 1);
                    if (expectedValuesRow == null)
                    {
                        int missingRowNum = rowIndex + 2; //+1 for 1-based, +1 for next row
                        throw new AssertFailedException("Missing expected values row for function '"
                                + targetFunctionName + " (row " + missingRowNum + ")");
                    }
                    switch (ProcessFunctionRow(evaluator, targetFunctionName, r, expectedValuesRow))
                    {
                        case Result.ALL_EVALUATIONS_SUCCEEDED: _functionSuccessCount++; break;
                        case Result.SOME_EVALUATIONS_FAILED: _functionFailureCount++; break;
                        case Result.NO_EVALUATIONS_FOUND: // do nothing
                            break;
                        default:
                            throw new Exception("unexpected result");
                    }
                }
                rowIndex += SS.NUMBER_OF_ROWS_PER_FUNCTION;
            }
        }

        /**
         * 
         * @return a constant from the local Result class denoting whether there were any evaluation
         * cases, and whether they all succeeded.
         */
        private int ProcessFunctionRow(HSSFFormulaEvaluator evaluator, String targetFunctionName,
                NPOI.SS.UserModel.IRow formulasRow, NPOI.SS.UserModel.IRow expectedValuesRow)
        {

            int result = Result.NO_EVALUATIONS_FOUND; // so far
            int endcolnum = formulasRow.LastCellNum;
            //evaluator.SetCurrentRow(formulasRow);

            // iterate across the row for all the evaluation cases
            for (short colnum = SS.COLUMN_INDEX_FIRST_TEST_VALUE; colnum < endcolnum; colnum++)
            {
                ICell c = formulasRow.GetCell(colnum);
                if (c == null || c.CellType != NPOI.SS.UserModel.CellType.FORMULA)
                {
                    continue;
                }

                NPOI.SS.UserModel.CellValue actualValue = evaluator.Evaluate(c);

                ICell expectedValueCell = GetExpectedValueCell(expectedValuesRow, colnum);
                try
                {
                    ConfirmExpectedResult("Function '" + targetFunctionName + "': Formula: " + c.CellFormula + " @ " + formulasRow.RowNum + ":" + colnum,
                            expectedValueCell, actualValue);
                    _evaluationSuccessCount++;
                    if (result != Result.SOME_EVALUATIONS_FAILED)
                    {
                        result = Result.ALL_EVALUATIONS_SUCCEEDED;
                    }
                }
                catch (AssertFailedException)
                {
                    _evaluationFailureCount++;
                    //printShortStackTrace(System.err, e);
                    result = Result.SOME_EVALUATIONS_FAILED;
                }
            }
            return result;
        }

        /**
         * @return <c>null</c> if cell is1 missing, empty or blank
         */
        private static String GetTargetFunctionName(NPOI.SS.UserModel.IRow r)
        {
            if (r == null)
            {
                Console.Error.WriteLine("Warning - given null row, can't figure out function name");
                return null;
            }
            ICell cell = r.GetCell(SS.COLUMN_INDEX_FUNCTION_NAME);
            if (cell == null)
            {
                Console.Error.WriteLine("Warning - NPOI.SS.UserModel.Row " + r.RowNum + " has no cell " + SS.COLUMN_INDEX_FUNCTION_NAME + ", can't figure out function name");
                return null;
            }
            if (cell.CellType == NPOI.SS.UserModel.CellType.BLANK)
            {
                return null;
            }
            if (cell.CellType == NPOI.SS.UserModel.CellType.STRING)
            {
                return cell.RichStringCellValue.String;
            }

            throw new AssertFailedException("Bad cell type for 'function name' column: ("
                    + cell.CellType + ") row (" + (r.RowNum + 1) + ")");
        }
    }
}
