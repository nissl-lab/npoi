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

namespace TestCases.SS.Formula.Atp
{
    using System;
    using System.Collections;
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.Atp;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.UserModel;
    using TestCases.HSSF;


    /**
     * Tests YearFracCalculator using Test-cases listed in a sample spreadsheet
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestYearFracCalculatorFromSpreadsheet
    {

        private static class SS
        {

            public static int BASIS_COLUMN = 1; // "B"
            public static int START_YEAR_COLUMN = 2; // "C"
            public static int END_YEAR_COLUMN = 5; // "F"
            public static int YEARFRAC_FORMULA_COLUMN = 11; // "L"
            public static int EXPECTED_RESULT_COLUMN = 13; // "N"
        }
        [Test]
        public void TestAll()
        {

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("yearfracExamples.xls");
            ISheet sheet = wb.GetSheetAt(0);
            HSSFFormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator(wb);
            int nSuccess = 0;
            int nFailures = 0;
            int nUnexpectedErrors = 0;
            IEnumerator rowIterator = sheet.GetRowEnumerator();
            while (rowIterator.MoveNext())
            {
                IRow row = (IRow)rowIterator.Current;

                ICell cell = row.GetCell(SS.YEARFRAC_FORMULA_COLUMN);
                if (cell == null || cell.CellType != CellType.Formula)
                {
                    continue;
                }
                try
                {
                    ProcessRow(row, cell, formulaEvaluator);
                    nSuccess++;
                }
                catch (SystemException e)
                {
                    nUnexpectedErrors++;
                    printshortStackTrace(System.Console.Error, e);
                }
                catch (AssertionException e)
                {
                    nFailures++;
                    printshortStackTrace(System.Console.Error, e);
                }
            }
            if (nUnexpectedErrors + nFailures > 0)
            {
                String msg = nFailures + " failures(s) and " + nUnexpectedErrors
                    + " unexpected errors(s) occurred. See stderr for details";
                throw new AssertionException(msg);
            }
            if (nSuccess < 1)
            {
                throw new Exception("No Test sample cases found");
            }
        }

        private static void ProcessRow(IRow row, ICell cell, HSSFFormulaEvaluator formulaEvaluator)
        {

            double startDate = MakeDate(row, SS.START_YEAR_COLUMN);
            double endDate = MakeDate(row, SS.END_YEAR_COLUMN);

            int basis = GetIntCell(row, SS.BASIS_COLUMN);

            double expectedValue = GetDoubleCell(row, SS.EXPECTED_RESULT_COLUMN);

            double actualValue;
            try
            {
                actualValue = YearFracCalculator.Calculate(startDate, endDate, basis);
            }
            catch (EvaluationException e)
            {
                throw e;
            }
            if (expectedValue != actualValue)
            {
                throw new Exception("Direct calculate failed - row " + (row.RowNum + 1) +
                        "excepted value " + expectedValue.ToString() + "actual value " + actualValue.ToString());
            }
            actualValue = formulaEvaluator.Evaluate(cell).NumberValue;
            if (expectedValue != actualValue)
            {

                throw new Exception("Formula Evaluate failed - row " + (row.RowNum + 1) +
                        "excepted value " + expectedValue.ToString() + "actual value " + actualValue.ToString());
            }
        }

        private static double MakeDate(IRow row, int yearColumn)
        {
            int year = GetIntCell(row, yearColumn + 0);
            int month = GetIntCell(row, yearColumn + 1);
            int day = GetIntCell(row, yearColumn + 2);
            DateTime d = new DateTime(year, month, day, 0, 0, 0, 0);
            return DateUtil.GetExcelDate(d);
        }

        private static int GetIntCell(IRow row, int colIx)
        {
            double dVal = GetDoubleCell(row, colIx);
            if (Math.Floor(dVal) != dVal)
            {
                throw new SystemException("Non integer value (" + dVal
                        + ") cell found at column " + (char)('A' + colIx));
            }
            return (int)dVal;
        }

        private static double GetDoubleCell(IRow row, int colIx)
        {
            ICell cell = row.GetCell(colIx);
            if (cell == null)
            {
                throw new SystemException("No cell found at column " + colIx);
            }
            double dVal = cell.NumericCellValue;
            return dVal;
        }

        /**
         * Useful to keep output concise when expecting many failures to be reported by this Test case
         * TODO - refactor duplicates in other Test~FromSpreadsheet classes
         */
        private static void printshortStackTrace(TextWriter ps, Exception e)
        {
            ps.Write(e.StackTrace);
            /*StackTraceElement[] stes = e.StackTrace;
		
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
            }
             * */
        }
    }

}