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


namespace TestCases.SS.Formula.Functions
{
    using System;
    using System.Text;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using TestCases.HSSF;
    using NPOI.SS.Formula.Eval;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Util;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;


    /**
     * Tests lookup functions (VLOOKUP, HLOOKUP, LOOKUP, MATCH) as loaded from a test data spReadsheet.<p/>
     * These tests have been separated from the common function and operator tests because the lookup
     * functions have more complex test cases and test data Setup.
     * 
     * Tests for bug fixes and specific/tricky behaviour can be found in the corresponding test class
     * (<tt>TestXxxx</tt>) of the tarGet (<tt>Xxxx</tt>) implementor, where execution can be observed
     *  more easily.
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestLookupFunctionsFromSpreadsheet
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

            /** Name of the test spReadsheet (found in the standard test data folder) */
            public static String FILENAME = "LookupFunctionsTestCaseData.xls";

            /** Name of the first sheet in the spReadsheet (contains comments) */
            public static String README_SHEET_NAME = "Read Me";


            /** Row (zero-based) in each sheet where the evaluation cases start.   */
            public static int START_TEST_CASES_ROW_INDEX = 4; // Row '5'
            /**  Index of the column that contains the function names */
            public static short COLUMN_INDEX_MARKER = 0; // Column 'A'
            public static short COLUMN_INDEX_EVALUATION = 1; // Column 'B'
            public static short COLUMN_INDEX_EXPECTED_RESULT = 2; // Column 'C'
            public static short COLUMN_ROW_COMMENT = 3; // Column 'D'

            /** Used to indicate when there are no more test cases on the current sheet   */
            public static String TEST_CASES_END_MARKER = "<end>";
            /** Used to indicate that the test on the current row should be ignored */
            public static String SKIP_CURRENT_TEST_CASE_MARKER = "<skip>";

        }

        // Note - multiple failures are aggregated before ending.  
        // If one or more functions fail, a single AssertFailedException is1 thrown at the end
        private int _sheetFailureCount;
        private int _sheetSuccessCount;
        private int _evaluationFailureCount;
        private int _evaluationSuccessCount;



        private static void ConfirmExpectedResult(String msg, ICell expected, NPOI.SS.UserModel.CellValue actual)
        {
            if (expected == null)
            {
                throw new AssertFailedException(msg + " - Bad Setup data expected value is1 null");
            }
            if (actual == null)
            {
                throw new AssertFailedException(msg + " - actual value was null");
            }
            if (expected.CellType == NPOI.SS.UserModel.CellType.ERROR)
            {
                ConfirmErrorResult(msg, Convert.ToInt32(expected.ErrorCellValue), actual);
                return;
            }
            if (actual.CellType == NPOI.SS.UserModel.CellType.ERROR)
            {
                throw UnexpectedError(msg, expected, actual.ErrorValue);
            }
            if (actual.CellType != expected.CellType)
            {
                WrongTypeError(msg, expected, actual);
            }


            switch (expected.CellType)
            {
                case NPOI.SS.UserModel.CellType.BOOLEAN:
                    Assert.AreEqual(expected.BooleanCellValue, actual.BooleanValue, msg);
                    break;
                case NPOI.SS.UserModel.CellType.FORMULA: // will never be used, since we will call method after formula evaluation
                    throw new AssertFailedException("Cannot expect formula as result of formula evaluation: " + msg);
                case NPOI.SS.UserModel.CellType.NUMERIC:
                    Assert.AreEqual(expected.NumericCellValue, actual.NumberValue, 0.0);
                    break;
                case NPOI.SS.UserModel.CellType.STRING:
                    Assert.AreEqual(expected.RichStringCellValue.String, actual.StringValue, msg);
                    break;
            }
        }


        private static AssertFailedException WrongTypeError(String msgPrefix, ICell expectedCell, NPOI.SS.UserModel.CellValue actualValue)
        {
            return new AssertFailedException(msgPrefix + " Result type mismatch. Evaluated result was "
                    + FormatValue(actualValue)
                    + " but the expected result was "
                    + FormatValue(expectedCell)
                    );
        }
        private static AssertFailedException UnexpectedError(String msgPrefix, ICell expected, int actualErrorCode)
        {
            return new AssertFailedException(msgPrefix + " Error code ("
                    + ErrorEval.GetText(actualErrorCode)
                    + ") was evaluated, but the expected result was "
                    + FormatValue(expected)
                    );
        }


        private static void ConfirmErrorResult(String msgPrefix, int expectedErrorCode, NPOI.SS.UserModel.CellValue actual)
        {
            if (actual.CellType != NPOI.SS.UserModel.CellType.ERROR)
            {
                throw new AssertFailedException(msgPrefix + " Expected cell error ("
                        + ErrorEval.GetText(expectedErrorCode) + ") but actual value was "
                        + FormatValue(actual));
            }
            if (expectedErrorCode != actual.ErrorValue)
            {
                throw new AssertFailedException(msgPrefix + " Expected cell error code ("
                        + ErrorEval.GetText(expectedErrorCode)
                        + ") but actual error code was ("
                        + ErrorEval.GetText(actual.ErrorValue)
                        + ")");
            }
        }


        private static String FormatValue(ICell expecedCell)
        {
            switch (expecedCell.CellType)
            {
                case NPOI.SS.UserModel.CellType.BLANK: return "<blank>";
                case NPOI.SS.UserModel.CellType.BOOLEAN: return expecedCell.BooleanCellValue.ToString();
                case NPOI.SS.UserModel.CellType.NUMERIC: return expecedCell.NumericCellValue.ToString();
                case NPOI.SS.UserModel.CellType.STRING: return expecedCell.RichStringCellValue.String;
            }
            throw new Exception("Unexpected cell type of expected value (" + expecedCell.CellType + ")");
        }
        private static String FormatValue(NPOI.SS.UserModel.CellValue actual)
        {
            switch (actual.CellType)
            {
                case NPOI.SS.UserModel.CellType.BLANK: return "<blank>";
                case NPOI.SS.UserModel.CellType.BOOLEAN: return actual.BooleanValue.ToString();
                case NPOI.SS.UserModel.CellType.NUMERIC: return actual.NumberValue.ToString();
                case NPOI.SS.UserModel.CellType.STRING: return actual.StringValue ;
            }
            throw new Exception("Unexpected cell type of evaluated value (" + actual.CellType + ")");
        }

        [TestInitialize]
        public void SetUp()
        {
            _sheetFailureCount = 0;
            _sheetSuccessCount = 0;
            _evaluationFailureCount = 0;
            _evaluationSuccessCount = 0;
        }
        [TestMethod]
        public void TestFunctionsFromTestSpreadsheet()
        {
            // This test methods depends on the american culture.
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook(SS.FILENAME);

            ConfirmReadMeSheet(workbook);
            int nSheets = workbook.NumberOfSheets;
            for (int i = 1; i < nSheets; i++)
            {
                int sheetResult = ProcessTestSheet(workbook, i, workbook.GetSheetName(i));
                switch (sheetResult)
                {
                    case Result.ALL_EVALUATIONS_SUCCEEDED: _sheetSuccessCount++; break;
                    case Result.SOME_EVALUATIONS_FAILED: _sheetFailureCount++; break;
                }
            }

            // Confirm results
            String successMsg = "There were "
                    + _sheetSuccessCount + " successful sheets(s) and "
                    + _evaluationSuccessCount + " function(s) without error";
            if (_sheetFailureCount > 0)
            {
                String msg = _sheetFailureCount + " sheets(s) failed with "
                + _evaluationFailureCount + " evaluation(s).  " + successMsg;
                throw new AssertFailedException(msg);
            }
            //if(false) { // normally no output for successful tests
            //    System.out.println(GetType().Name + ": " + successMsg);
            //}
        }

        private int ProcessTestSheet(HSSFWorkbook workbook, int sheetIndex, String sheetName)
        {
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(sheetIndex);
            HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(sheet, workbook);
            int maxRows = sheet.LastRowNum + 1;
            int result = Result.NO_EVALUATIONS_FOUND; // so far

            String currentGroupComment = null;
            for (int rowIndex = SS.START_TEST_CASES_ROW_INDEX; rowIndex < maxRows; rowIndex++)
            {
                IRow r = sheet.GetRow(rowIndex);
                String newMarkerValue = GetMarkerColumnValue(r);
                if (r == null)
                {
                    continue;
                }
                if (SS.TEST_CASES_END_MARKER.Equals(newMarkerValue, StringComparison.InvariantCultureIgnoreCase))
                {
                    // normal exit point
                    return result;
                }
                if (SS.SKIP_CURRENT_TEST_CASE_MARKER.Equals(newMarkerValue, StringComparison.InvariantCultureIgnoreCase))
                {
                    // currently disabled test case row
                    continue;
                }
                if (newMarkerValue != null)
                {
                    currentGroupComment = newMarkerValue;
                }
                ICell c = r.GetCell(SS.COLUMN_INDEX_EVALUATION);
                if (c == null || c.CellType != NPOI.SS.UserModel.CellType.FORMULA)
                {
                    continue;
                }
                //evaluator.SetCurrentRow(r);
                NPOI.SS.UserModel.CellValue actualValue = evaluator.Evaluate(c);
                ICell expectedValueCell = r.GetCell(SS.COLUMN_INDEX_EXPECTED_RESULT);
                String rowComment = GetRowCommentColumnValue(r);

                String msgPrefix = FormatTestCaseDetails(sheetName, r.RowNum, c, currentGroupComment, rowComment);
                try
                {
                    ConfirmExpectedResult(msgPrefix, expectedValueCell, actualValue);
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
                catch (Exception)
                {
                    _evaluationFailureCount++;
                    //printShortStackTrace(System.err, e);
                    result = Result.SOME_EVALUATIONS_FAILED;
                }
            }
            throw new Exception("Missing end marker '" + SS.TEST_CASES_END_MARKER
                    + "' on sheet '" + sheetName + "'");

        }


        private static String FormatTestCaseDetails(String sheetName, int rowNum, ICell c, String currentGroupComment,
                String rowComment)
        {

            StringBuilder sb = new StringBuilder();
            CellReference cr = new CellReference(sheetName, rowNum, c.ColumnIndex, false, false);
            sb.Append(cr.FormatAsString());
            sb.Append(" {=").Append(c.CellFormula).Append("}");

            if (currentGroupComment != null)
            {
                sb.Append(" '");
                sb.Append(currentGroupComment);
                if (rowComment != null)
                {
                    sb.Append(" - ");
                    sb.Append(rowComment);
                }
                sb.Append("' ");
            }
            else
            {
                if (rowComment != null)
                {
                    sb.Append(" '");
                    sb.Append(rowComment);
                    sb.Append("' ");
                }
            }

            return sb.ToString();
        }

        /**
         * Asserts that the 'Read me' comment page exists, and has this class' name in one of the 
         * cells.  This back-link is1 to make it easy to find this class if a Reader encounters the 
         * spReadsheet first.
         */
        private void ConfirmReadMeSheet(HSSFWorkbook workbook)
        {
            String firstSheetName = workbook.GetSheetName(0);
            if (!firstSheetName.Equals(SS.README_SHEET_NAME,StringComparison.InvariantCultureIgnoreCase))
            {
                throw new Exception("First sheet's name was '" + firstSheetName + "' but expected '" + SS.README_SHEET_NAME + "'");
            }
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(0);
            String specifiedClassName = sheet.GetRow(2).GetCell((short)0).RichStringCellValue.String;
            Assert.AreEqual(GetType().Name, specifiedClassName.Substring(specifiedClassName.LastIndexOf(".") + 1), "Test class name in spreadsheet comment");

        }

        private static String GetRowCommentColumnValue(IRow r)
        {
            return GetCellTextValue(r, SS.COLUMN_ROW_COMMENT, "row comment");
        }

        private static String GetMarkerColumnValue(IRow r)
        {
            return GetCellTextValue(r, SS.COLUMN_INDEX_MARKER, "marker");
        }

        /**
         * @return <c>null</c> if cell is1 missing, empty or blank
         */
        private static String GetCellTextValue(IRow r, int colIndex, String columnName)
        {
            if (r == null)
            {
                return null;
            }
            ICell cell = r.GetCell((short)colIndex);
            if (cell == null)
            {
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

            throw new Exception("Bad cell type for '" + columnName + "' column: ("
                    + cell.CellType + ") row (" + (r.RowNum + 1) + ")");
        }
    }
}