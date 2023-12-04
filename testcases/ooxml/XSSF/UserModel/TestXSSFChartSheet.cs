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

using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Util;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace TestCases.XSSF.UserModel
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

            CT_Chartsheet ctChartsheet = ((XSSFChartSheet)wb.GetSheetAt(2)).GetCTChartsheet();
            Assert.IsNotNull(ctChartsheet);
        }
        [Test]
        public void TestGetAccessors()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chart_sheet.xlsx");
            XSSFChartSheet sheet = (XSSFChartSheet)wb.GetSheetAt(2);

            Assert.IsFalse(sheet.GetEnumerator().MoveNext(),
                "Row iterator for charts sheets should return zero rows");

            //access to a arbitrary row
            Assert.IsNull(sheet.GetRow(1));

            //some basic get* accessors
            Assert.AreEqual(0, sheet.NumberOfComments);
            Assert.AreEqual(0, sheet.NumHyperlinks);
            Assert.AreEqual(0, sheet.NumMergedRegions);
            Assert.IsNull(sheet.ActiveCell);
            Assert.IsTrue(sheet.Autobreaks);
            Assert.IsNull(sheet.GetCellComment(new CellAddress(0, 0)));
            Assert.IsNull(sheet.GetCellComment(new CellAddress(0, 0)));
            Assert.AreEqual(0, sheet.ColumnBreaks.Length);
            Assert.IsTrue(sheet.RowSumsBelow);

            Assert.IsNotNull(sheet.CreateDrawingPatriarch());
            Assert.IsNotNull(sheet.GetDrawingPatriarch());
            Assert.IsNotNull(sheet.GetCTChartsheet());
        }

        [Test]
        public void TestGetCharts()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chart_sheet.xlsx");

            XSSFSheet ns = (XSSFSheet)wb.GetSheetAt(0);
            XSSFChartSheet cs = (XSSFChartSheet)wb.GetSheetAt(2);

            Assert.AreEqual(0, (ns.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);
            Assert.AreEqual(1, (cs.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);

            XSSFChart chart = (cs.CreateDrawingPatriarch() as XSSFDrawing).GetCharts()[0];
            Assert.IsNull(chart.Title);
        }
    }
}