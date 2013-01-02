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

using System;
using NUnit.Framework;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
namespace TestCases.SS.Formula.Functions
{


    /**
     * Tests for {@link TimeFunc}
     *
     * @author @author Steven Butler (sebutler @ gmail dot com)
     */
    [TestFixture]
    public class TestTime
    {

        private static int SECONDS_PER_MINUTE = 60;
        private static int SECONDS_PER_HOUR = 60 * SECONDS_PER_MINUTE;
        private static double SECONDS_PER_DAY = 24 * SECONDS_PER_HOUR;
        private ICell cell11;
        private HSSFFormulaEvaluator Evaluator;
        private HSSFWorkbook wb;
        private DataFormatter form;
        private ICellStyle style;
        [SetUp]
        public void SetUp()
        {
            wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("new sheet");
            style = wb.CreateCellStyle();
            IDataFormat fmt = wb.CreateDataFormat();
            style.DataFormat=(fmt.GetFormat("hh:mm:ss"));

            cell11 = sheet.CreateRow(0).CreateCell(0);
            form = new DataFormatter();

            Evaluator = new HSSFFormulaEvaluator(wb);
        }
        [Test]
        public void TestSomeArgumentsMissing()
        {
            Confirm("00:00:00", "TIME(, 0, 0)");
            Confirm("12:00:00", "TIME(12, , )");
        }
        [Test]
        public void TestValid()
        {
            Confirm("00:00:01", 0, 0, 1);
            Confirm("00:01:00", 0, 1, 0);

            Confirm("00:00:00", 0, 0, 0);

            Confirm("01:00:00", 1, 0, 0);
            Confirm("12:00:00", 12, 0, 0);
            Confirm("23:00:00", 23, 0, 0);
            Confirm("00:00:00", 24, 0, 0);
            Confirm("01:00:00", 25, 0, 0);
            Confirm("00:00:00", 48, 0, 0);
            Confirm("06:00:00", 6, 0, 0);
            Confirm("06:01:00", 6, 1, 0);
            Confirm("06:30:00", 6, 30, 0);

            Confirm("06:59:00", 6, 59, 0);
            Confirm("07:00:00", 6, 60, 0);
            Confirm("07:01:00", 6, 61, 0);
            Confirm("08:00:00", 6, 120, 0);
            Confirm("06:00:00", 6, 1440, 0);
            Confirm("18:49:00", 18, 49, 0);
            Confirm("18:49:01", 18, 49, 1);
            Confirm("18:49:30", 18, 49, 30);
            Confirm("18:49:59", 18, 49, 59);
            Confirm("18:50:00", 18, 49, 60);
            Confirm("18:50:01", 18, 49, 61);
            Confirm("18:50:59", 18, 49, 119);
            Confirm("18:51:00", 18, 49, 120);
            Confirm("03:55:07", 18, 49, 32767);
            Confirm("12:08:01", 18, 32767, 61);
            Confirm("07:50:01", 32767, 49, 61);
        }
        private void Confirm(String expectedTimeStr, int inH, int inM, int inS)
        {
            Confirm(expectedTimeStr, "TIME(" + inH + "," + inM + "," + inS + ")");
        }

        private void Confirm(String expectedTimeStr, String formulaText)
        {
            //		Console.WriteLine("=" + formulaText);
            //String[] parts = Pattern.compile(":").split(expectedTimeStr);
            string[] parts = expectedTimeStr.Split(":".ToCharArray());
            int expH = Int32.Parse(parts[0]);
            int expM = Int32.Parse(parts[1]);
            int expS = Int32.Parse(parts[2]);

            double expectedValue = (expH * SECONDS_PER_HOUR + expM * SECONDS_PER_MINUTE + expS) / SECONDS_PER_DAY;

            cell11.CellFormula=(formulaText);
            cell11.CellStyle=(style);
            Evaluator.ClearAllCachedResultValues();

            double actualValue = Evaluator.Evaluate(cell11).NumberValue;
            Assert.AreEqual(expectedValue, actualValue, 0.0);

            String actualText = form.FormatCellValue(cell11, Evaluator);
            Assert.AreEqual(expectedTimeStr, actualText);
        }
    }
}

