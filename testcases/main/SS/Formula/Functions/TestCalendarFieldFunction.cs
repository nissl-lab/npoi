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

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NUnit.Framework;
    using System;

    /**
     * Test for YEAR / MONTH / DAY / HOUR / MINUTE / SECOND
     */
    [TestFixture]
    public class TestCalendarFieldFunction
    {

        private ICell cell11;
        private HSSFFormulaEvaluator Evaluator;
        [SetUp]
        public void SetUp()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("new sheet");
            cell11 = sheet.CreateRow(0).CreateCell(0);
            cell11.SetCellType(CellType.Formula);
            Evaluator = new HSSFFormulaEvaluator(wb);
        }
        [Test]
        public void TestValid()
        {
            Confirm("YEAR(2.26)", 1900);
            Confirm("MONTH(2.26)", 1);
            Confirm("DAY(2.26)", 2);
            Confirm("HOUR(2.26)", 6);
            Confirm("MINUTE(2.26)", 14);
            Confirm("SECOND(2.26)", 24);

            Confirm("YEAR(40627.4860417)", 2011);
            Confirm("MONTH(40627.4860417)", 3);
            Confirm("DAY(40627.4860417)", 25);
            Confirm("HOUR(40627.4860417)", 11);
            Confirm("MINUTE(40627.4860417)", 39);
            Confirm("SECOND(40627.4860417)", 54);
        }
        [Test]
        public void TestRounding()
        {
            // 41484.999994200 = 23:59:59,499
            // 41484.9999942129 = 23:59:59,500  (but sub-milliseconds are below 0.5 (0.49999453965575), XLS-second results in 59)
            // 41484.9999942130 = 23:59:59,500  (sub-milliseconds are 0.50000334065408, XLS-second results in 00)

            Confirm("DAY(41484.999994200)", 29);
            Confirm("SECOND(41484.999994200)", 59);

            Confirm("DAY(41484.9999942129)", 29);
            Confirm("HOUR(41484.9999942129)", 23);
            Confirm("MINUTE(41484.9999942129)", 59);
            Confirm("SECOND(41484.9999942129)", 59);

            Confirm("DAY(41484.9999942130)", 30);
            Confirm("HOUR(41484.9999942130)", 0);
            Confirm("MINUTE(41484.9999942130)", 0);
            Confirm("SECOND(41484.9999942130)", 0);
        }
        [Test]
        public void TestDaylightSaving()
        {
            Confirm("HOUR(41364.08263888890000)", 1);		// 31.03.2013 01:59:00,000
            Confirm("HOUR(41364.08333333330000)", 2);		// 31.03.2013 02:00:00,000 (this time does not exist in TZ CET, but EXCEL does not care)
            Confirm("HOUR(41364.08402777780000)", 2);		// 31.03.2013 02:01:00,000
            Confirm("HOUR(41364.12430555560000)", 2);		// 31.03.2013 02:59:00,000
            Confirm("HOUR(41364.12500000000000)", 3);		// 31.03.2013 03:00:00,000
        }
        [Test]
        public void TestBugDate()
        {
            Confirm("YEAR(0.0)", 1900);
            Confirm("MONTH(0.0)", 1);
            Confirm("DAY(0.0)", 0);

            Confirm("YEAR(0.26)", 1900);
            Confirm("MONTH(0.26)", 1);
            Confirm("DAY(0.26)", 0);
            Confirm("HOUR(0.26)", 6);
            Confirm("MINUTE(0.26)", 14);
            Confirm("SECOND(0.26)", 24);
        }

        private void Confirm(String formulaText, double expectedResult)
        {
            cell11.CellFormula=(formulaText);
            Evaluator.ClearAllCachedResultValues();
            CellValue cv = Evaluator.Evaluate(cell11);
            if (cv.CellType != CellType.Numeric)
            {
                throw new AssertionException("Wrong result type: " + cv.FormatAsString());
            }
            double actualValue = cv.NumberValue;
            Assert.AreEqual(expectedResult, actualValue, 0);
        }
    }

}