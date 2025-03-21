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
using NUnit.Framework;using NUnit.Framework.Legacy;

namespace TestCases.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFChartSheet
    {
        [Test]
        public void TestXSSFFactory()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chart_sheet.xlsx");
            ClassicAssert.AreEqual(4, wb.NumberOfSheets);

            //the third sheet is of type 'chartsheet'
            ClassicAssert.AreEqual("Chart1", wb.GetSheetName(2));
            ClassicAssert.IsTrue(wb.GetSheetAt(2) is XSSFChartSheet);
            ClassicAssert.AreEqual("Chart1", wb.GetSheetAt(2).SheetName);

            CT_Chartsheet ctChartsheet = ((XSSFChartSheet)wb.GetSheetAt(2)).GetCTChartsheet();
            ClassicAssert.IsNotNull(ctChartsheet);
        }
        [Test]
        public void TestGetAccessors()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chart_sheet.xlsx");
            XSSFChartSheet sheet = (XSSFChartSheet)wb.GetSheetAt(2);

            ClassicAssert.IsFalse(sheet.GetEnumerator().MoveNext(),
                "Row iterator for charts sheets should return zero rows");

            //access to a arbitrary row
            ClassicAssert.IsNull(sheet.GetRow(1));

            //some basic get* accessors
            ClassicAssert.AreEqual(0, sheet.NumberOfComments);
            ClassicAssert.AreEqual(0, sheet.NumHyperlinks);
            ClassicAssert.AreEqual(0, sheet.NumMergedRegions);
            ClassicAssert.IsNull(sheet.ActiveCell);
            ClassicAssert.IsTrue(sheet.Autobreaks);
            ClassicAssert.IsNull(sheet.GetCellComment(new CellAddress(0, 0)));
            ClassicAssert.IsNull(sheet.GetCellComment(new CellAddress(0, 0)));
            ClassicAssert.AreEqual(0, sheet.ColumnBreaks.Length);
            ClassicAssert.IsTrue(sheet.RowSumsBelow);

            ClassicAssert.IsNotNull(sheet.CreateDrawingPatriarch());
            ClassicAssert.IsNotNull(sheet.GetDrawingPatriarch());
            ClassicAssert.IsNotNull(sheet.GetCTChartsheet());
        }

        [Test]
        public void TestGetCharts()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("chart_sheet.xlsx");

            XSSFSheet ns = (XSSFSheet)wb.GetSheetAt(0);
            XSSFChartSheet cs = (XSSFChartSheet)wb.GetSheetAt(2);

            ClassicAssert.AreEqual(0, (ns.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);
            ClassicAssert.AreEqual(1, (cs.CreateDrawingPatriarch() as XSSFDrawing).GetCharts().Count);

            XSSFChart chart = (cs.CreateDrawingPatriarch() as XSSFDrawing).GetCharts()[0];
            ClassicAssert.IsNull(chart.Title);
        }
    }
}