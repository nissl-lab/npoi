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

    /**
     * Tests INDEX() as loaded from a Test data spreadsheet.<p/>
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestIndexFunctionFromSpreadsheet
    {

        private static class Result
        {
            public static int SOME_EVALUATIONS_FAILED = -1;
            public static int ALL_EVALUATIONS_SUCCEEDED = +1;
            public static int NO_EVALUATIONS_FOUND = 0;
        }

        /**
         * This class defines constants for navigating around the Test data spreadsheet used for these Tests.
         */
        private static class SS
        {

            /** Name of the Test spreadsheet (found in the standard Test data folder) */
            public static String FILENAME = "IndexFunctionTestCaseData.xls";

            public static int COLUMN_INDEX_EVALUATION = 2; // Column 'C'
            public static int COLUMN_INDEX_EXPECTED_RESULT = 3; // Column 'D'

        }

        // Note - multiple failures are aggregated before ending.
        // If one or more functions fail, a single AssertionException is thrown at the end
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
            _EvaluationFailureCount = 0;
            _EvaluationSuccessCount = 0;
        }
        [Test]
        public void TestFunctionsFromTestSpreadsheet()
        {
            HSSFWorkbook workbook = HSSFTestDataSamples.OpenSampleWorkbook(SS.FILENAME);

            ProcessTestSheet(workbook, workbook.GetSheetName(0));

            // confirm results
            String successMsg = "There were "
                    + _EvaluationSuccessCount + " function(s) without error";
            if (_EvaluationFailureCount > 0)
            {
                String msg = _EvaluationFailureCount + " Evaluation(s) failed.  " + successMsg;
                throw new AssertionException(msg);
            }
#if !HIDE_UNREACHABLE_CODE
            if (false)
            { // normally no output for successful Tests
                Console.WriteLine(this.GetType().Name + ": " + successMsg);
            }
#endif
        }

        private void ProcessTestSheet(HSSFWorkbook workbook, String sheetName)
        {
            ISheet sheet = workbook.GetSheetAt(0);
            HSSFFormulaEvaluator Evaluator = new HSSFFormulaEvaluator(workbook);
            int maxRows = sheet.LastRowNum + 1;
            int result = Result.NO_EVALUATIONS_FOUND; // so far

            for (int rowIndex = 0; rowIndex < maxRows; rowIndex++)
            {
                IRow r = sheet.GetRow(rowIndex);
                if (r == null)
                {
                    continue;
                }
                ICell c = r.GetCell(SS.COLUMN_INDEX_EVALUATION);
                if (c == null || c.CellType != CellType.Formula)
                {
                    continue;
                }
                ICell expectedValueCell = r.GetCell(SS.COLUMN_INDEX_EXPECTED_RESULT);

                String msgPrefix = formatTestCaseDetails(sheetName, r.RowNum, c);
                try
                {
                    CellValue actualValue = Evaluator.Evaluate(c);
                    ConfirmExpectedResult(msgPrefix, expectedValueCell, actualValue);
                    _EvaluationSuccessCount++;
                    if (result != Result.SOME_EVALUATIONS_FAILED)
                    {
                        result = Result.ALL_EVALUATIONS_SUCCEEDED;
                    }
                }
                catch (SystemException e)
                {
                    _EvaluationFailureCount++;
                    printshortStackTrace(System.Console.Error, e, msgPrefix);
                    result = Result.SOME_EVALUATIONS_FAILED;
                }
                catch (AssertionException e)
                {
                    _EvaluationFailureCount++;
                    printshortStackTrace(System.Console.Error, e, msgPrefix);
                    result = Result.SOME_EVALUATIONS_FAILED;
                }
            }
        }


        private static String formatTestCaseDetails(String sheetName, int rowIndex, ICell c)
        {

            StringBuilder sb = new StringBuilder();
            CellReference cr = new CellReference(sheetName, rowIndex, c.ColumnIndex, false, false);
            sb.Append(cr.FormatAsString());
            sb.Append(" [formula: ").Append(c.CellFormula).Append(" ]");
            return sb.ToString();
        }

        /**
         * Useful to keep output concise when expecting many failures to be reported by this Test case
         */
        private static void printshortStackTrace(TextWriter ps, Exception e, String msgPrefix)
        {

            ps.WriteLine("Problem with " + msgPrefix);
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
    }

}
