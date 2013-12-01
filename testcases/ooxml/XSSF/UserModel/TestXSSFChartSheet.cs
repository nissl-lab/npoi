/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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

using NPOI.SS.UserModel;
using NUnit.Framework;

namespace NPOI.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFChartSheet
    {
        [Test]
        public void TestXSSFFactory()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chart_sheet.xlsx");
            Assert.AreEqual(4, wb.NumberOfSheets);

            //the third sheet is of type 'chartsheet'
            Assert.AreEqual("Chart1", wb.GetSheetName(2));
            Assert.IsTrue(wb.GetSheetAt(2) is XSSFChartSheet);
            Assert.AreEqual("Chart1", wb.GetSheetAt(2).SheetName);

        }
        [Test]
        public void TestGetAccessors()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chart_sheet.xlsx");
            XSSFChartSheet sheet = (XSSFChartSheet) wb.GetSheetAt(2);

            //foreach (IRow row in sheet)
            //{
            //    fail("Row iterator for charts sheets should return zero rows");
            //}

            //access to a arbitrary row
            Assert.AreEqual(null, sheet.GetRow(1));

            //some basic get* accessors
            Assert.AreEqual(0, sheet.NumberOfComments);
            Assert.AreEqual(0, sheet.NumHyperlinks);
            Assert.AreEqual(0, sheet.NumMergedRegions);
            Assert.AreEqual(null, sheet.ActiveCell);
            Assert.AreEqual(true, sheet.Autobreaks);
            Assert.AreEqual(null, sheet.GetCellComment(0, 0));
            Assert.AreEqual(0, sheet.ColumnBreaks.Length);
            Assert.AreEqual(true, sheet.RowSumsBelow);
        }

        [Test]
        public void TestGetCharts()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chart_sheet.xlsx");

            XSSFSheet ns = (XSSFSheet) wb.GetSheetAt(0);
            XSSFChartSheet cs = (XSSFChartSheet) wb.GetSheetAt(2);

            Assert.AreEqual(0, (ns.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);
            Assert.AreEqual(1, (cs.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);

            XSSFChart chart = (cs.CreateDrawingPatriarch() as XSSFDrawing).GetCharts()[0];
            Assert.AreEqual(null, chart.Title);
        }
    }
}