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

namespace TestCases.SS.Formula.Functions
{

    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using System;
    using NUnit.Framework;
    using TestCases.HSSF;
    using NPOI.HSSF.UserModel;
    using System.Text;
    using NPOI.SS.Util;
    using System.IO;
    using NPOI.Util;

    /**
     * Tests INDEX() as loaded from a Test data spreadsheet.<p/>
     *
     * @author Josh Micich
     */
    [TestFixture]
    public abstract class BaseTestFunctionsFromSpreadsheet
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
            /** Name of the first sheet in the spreadsheet (contains comments) */
            public const String README_SHEET_NAME = "Read Me";

            /** Row (zero-based) in each sheet where the evaluation cases start.   */
            public const int START_TEST_CASES_ROW_INDEX = 4; // Row '5'
            /**  Index of the column that contains the function names */
            public const int COLUMN_INDEX_MARKER = 0; // Column 'A'
            public const int COLUMN_INDEX_EVALUATION = 1; // Column 'B'
            public const int COLUMN_INDEX_EXPECTED_RESULT = 2; // Column 'C'
            public const int COLUMN_ROW_COMMENT = 3; // Column 'D'
            /** Used to indicate when there are no more test cases on the current sheet   */
            public const String TEST_CASES_END_MARKER = "<end>";
            /** Used to indicate that the test on the current row should be ignored */
            public const String SKIP_CURRENT_TEST_CASE_MARKER = "<skip>";
        }
        private int _sheetFailureCount;
        private int _sheetSuccessCount;
        // Note - multiple failures are aggregated before ending.
        // If one or more functions fail, a single AssertionException is thrown at the end
        private int _evaluationFailureCount;
        private int _evaluationSuccessCount;



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
            if (expected.CellType == CellType.Error)
            {
                ConfirmErrorResult(msg, expected.ErrorCellValue, actual);
                return;
            }
            if (actual.CellType == CellType.Error)
            {
                throw unexpectedError(msg, expected, actual.ErrorValue);
            }
            if (actual.CellType != expected.CellType)
            {
                throw wrongTypeError(msg, expected, actual);
            }


            switch (expected.CellType)
            {
                case CellType.Boolean:
                    Assert.AreEqual(expected.BooleanCellValue, actual.BooleanValue, msg);
                    break;
                case CellType.Formula: // will never be used, since we will call method After formula Evaluation
                    throw new AssertionException("Cannot expect formula as result of formula Evaluation: " + msg);
                case CellType.Numeric:
                    Assert.AreEqual(expected.NumericCellValue, actual.NumberValue, 0.0, msg);
                    break;
                case CellType.String:
                    Assert.AreEqual(expected.RichStringCellValue.String, actual.StringValue, msg);
                    break;
            }
        }


        private static AssertionException wrongTypeError(String msgPrefix, ICell expectedCell, CellValue actualValue)
        {
            return new AssertionException(msgPrefix + " Result type mismatch. Evaluated result was "
                    + actualValue.FormatAsString()
                    + " but the expected result was "
                    + formatValue(expectedCell)
                    );
        }
        private static AssertionException unexpectedError(String msgPrefix, ICell expected, int actualErrorCode)
        {
            return new AssertionException(msgPrefix + " Error code ("
                    + ErrorEval.GetText(actualErrorCode)
                    + ") was Evaluated, but the expected result was "
                    + formatValue(expected)
                    );
        }


        private static void ConfirmErrorResult(String msgPrefix, int expectedErrorCode, CellValue actual)
        {
            if (actual.CellType != CellType.Error)
            {
                throw new AssertionException(msgPrefix + " Expected cell error ("
                        + ErrorEval.GetText(expectedErrorCode) + ") but actual value was "
                        + actual.FormatAsString());
            }
            if (expectedErrorCode != actual.ErrorValue)
            {
                throw new AssertionException(msgPrefix + " Expected cell error code ("
                        + ErrorEval.GetText(expectedErrorCode)
                        + ") but actual error code was ("
                        + ErrorEval.GetText(actual.ErrorValue)
                        + ")");
            }
        }


        private static String formatValue(ICell expecedCell)
        {
            switch (expecedCell.CellType)
            {
                case CellType.Blank: return "<blank>";
                case CellType.Boolean: return expecedCell.BooleanCellValue.ToString();
                case CellType.Numeric: return expecedCell.NumericCellValue.ToString();
                case CellType.String: return expecedCell.RichStringCellValue.String;
            }
            throw new Exception("Unexpected cell type of expected value (" + expecedCell.CellType + ")");
        }

        [SetUp]
        public void SetUp()
        {
            _sheetFailureCount = 0;
            _sheetSuccessCount = 0;
            _evaluationFailureCount = 0;
            _evaluationSuccessCount = 0;
        }
        [Test]
        public void TestFunctionsFromTestSpreadsheet()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook(this.Filename);
            confirmReadMeSheet(workbook);
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

            // confirm results
            String successMsg = "There were "
                    + _sheetSuccessCount + " successful sheets(s) and "
                    + _evaluationSuccessCount + " function(s) without error";
            if (_sheetFailureCount > 0)
            {
                String msg = _sheetFailureCount + " sheets(s) failed with "
                        + _evaluationFailureCount + " evaluation(s).  " + successMsg;
                throw new AssertionException(msg);
            }
#if !HIDE_UNREACHABLE_CODE
            if (false)
            { // normally no output for successful Tests
                Console.WriteLine(this.GetType().Name + ": " + successMsg);
            }
#endif
        }

        private int ProcessTestSheet(HSSFWorkbook workbook, int sheetIndex, String sheetName)
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            HSSFFormulaEvaluator Evaluator = new HSSFFormulaEvaluator(workbook);
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
                if (c == null || c.CellType != CellType.Formula)
                {
                    continue;
                }
                ICell expectedValueCell = r.GetCell(SS.COLUMN_INDEX_EXPECTED_RESULT);
                String rowComment = GetRowCommentColumnValue(r);

                String msgPrefix = formatTestCaseDetails(this.Filename,sheetName, r.RowNum, c, currentGroupComment, rowComment);
                try
                {
                    CellValue actualValue = Evaluator.Evaluate(c);
                    ConfirmExpectedResult(msgPrefix, expectedValueCell, actualValue);
                    _evaluationSuccessCount++;
                    if (result != Result.SOME_EVALUATIONS_FAILED)
                    {
                        result = Result.ALL_EVALUATIONS_SUCCEEDED;
                    }
                }
                catch (RuntimeException e)
                {
                    _evaluationFailureCount++;
                    printshortStackTrace(System.Console.Error, e);
                    result = Result.SOME_EVALUATIONS_FAILED;
                }
                catch (AssertionException e)
                {
                    _evaluationFailureCount++;
                    printshortStackTrace(System.Console.Error, e);
                    result = Result.SOME_EVALUATIONS_FAILED;
                }
            }
            throw new Exception("Missing end marker '" + SS.TEST_CASES_END_MARKER
        + "' on sheet '" + sheetName + "'");
        }

        protected abstract String Filename
        {
            get;
        }
        private static String formatTestCaseDetails(String filename, String sheetName, int rowIndex, ICell c, String currentGroupComment,
                                                String rowComment)
        {

            StringBuilder sb = new StringBuilder();
            sb.Append("In ").Append(filename).Append(" ");

            CellReference cr = new CellReference(sheetName, rowIndex, c.ColumnIndex, false, false);
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
 * Asserts that the 'read me' comment page exists, and has this class' name in one of the
 * cells.  This back-link is to make it easy to find this class if a reader encounters the
 * spreadsheet first.
 */
        private void confirmReadMeSheet(HSSFWorkbook workbook)
        {
            String firstSheetName = workbook.GetSheetName(0);
            if (!firstSheetName.Equals(SS.README_SHEET_NAME, StringComparison.InvariantCultureIgnoreCase))
            {
                throw new Exception("First sheet's name was '" + firstSheetName + "' but expected '" + SS.README_SHEET_NAME + "'");
            }
            //ISheet sheet = workbook.GetSheetAt(0);
            //String specifiedClassName = sheet.GetRow(2).GetCell(0).RichStringCellValue.String;
            //Assert.AreEqual(this.GetType().FullName, specifiedClassName, "Test class name in spreadsheet comment");
        }

        /**
         * Useful to keep output concise when expecting many failures to be reported by this Test case
         */
        private static void printshortStackTrace(TextWriter ps, Exception e)
        {
            ps.WriteLine(e.Message);
            ps.WriteLine(e.StackTrace);
            //StackTraceElement[] stes = e.GetStackTrace();

            //int startIx = 0;
            //// skip any top frames inside junit.framework.Assert
            //while(startIx<stes.Length) {
            //    if(!stes[startIx].GetClassName().Equals(typeof(Assert).Name)) {
            //        break;
            //    }
            //    startIx++;
            //}
            //// skip bottom frames (part of junit framework)
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
            //endIx -= 4; // skip 4 frames of reflection invocation
            //ps.println(e.ToString());
            //for(int i=startIx; i<endIx; i++) {
            //    ps.println("\tat " + stes[i].ToString());
            //}
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
         * @return <code>null</code> if cell is missing, empty or blank
         */
        private static String GetCellTextValue(IRow r, int colIndex, String columnName)
        {
            if (r == null)
            {
                return null;
            }
            ICell cell = r.GetCell(colIndex);
            if (cell == null)
            {
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

            throw new Exception("Bad cell type for '" + columnName + "' column: ("
                    + cell.CellType + ") row (" + (r.RowNum + 1) + ")");
        }
    }

}
