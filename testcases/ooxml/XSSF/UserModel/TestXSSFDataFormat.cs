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

using TestCases.SS.UserModel;
using NUnit.Framework;
using NPOI.SS.UserModel;
using NPOI.XSSF;
using NPOI.XSSF.UserModel;

namespace TestCases.XSSF.UserModel
{

    /**
     * Tests for {@link XSSFDataFormat}
     */
    [TestFixture]
    public class TestXSSFDataFormat : BaseTestDataFormat
    {

        public TestXSSFDataFormat()
            : base(XSSFITestDataProvider.instance)
        {

        }

        /**
         * [Bug 49928] formatCellValue returns incorrect value for \u00a3 formatted cells
         */
        [Test]
        public override void Test49928()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("49928.xlsx");
            DoTest49928Core(wb);
            IDataFormat dataFormat = wb.CreateDataFormat();

            // As of 2015-12-27, there is no way to override a built-in number format with POI XSSFWorkbook
            // 49928.xlsx has been saved with a poundFmt that overrides the default value (dollar)
            short poundFmtIdx = wb.GetSheetAt(0).GetRow(0).GetCell(0).CellStyle.DataFormat;
            Assert.AreEqual(poundFmtIdx, dataFormat.GetFormat(poundFmt));

            // now create a custom format with Pound (\u00a3)

            string customFmt = "\u00a3##.00[Yellow]";
            AssertNotBuiltInFormat(customFmt);
            short customFmtIdx = dataFormat.GetFormat(customFmt);

            Assert.IsTrue(customFmtIdx >= BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX);
            Assert.AreEqual(customFmt, dataFormat.GetFormat(customFmtIdx));

            wb.Close();
        }

        /**
         * [Bug 58532] Handle formats that go numnum, numK, numM etc 
         */
        [Test]
        public void Test58532()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("FormatKM.xlsx");
            DoTest58532Core(wb);
            wb.Close();
        }

        /**
         * [Bug 58778] Built-in number formats can be overridden with XSSFDataFormat.putFormat(int id, String fmt)
         */
        [Test]
        public void Test58778()
        {
            XSSFWorkbook wb1 = new XSSFWorkbook();
            ICell cell = wb1.CreateSheet("bug58778").CreateRow(0).CreateCell(0);
            cell.SetCellValue(5.25);
            ICellStyle style = wb1.CreateCellStyle();

            XSSFDataFormat dataFormat = wb1.CreateDataFormat() as XSSFDataFormat;

            short poundFmtIdx = 6;
            dataFormat.PutFormat(poundFmtIdx, poundFmt);
            style.DataFormat = (poundFmtIdx);
            cell.CellStyle = style;
            // Cell should appear as "<poundsymbol>5"

            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutCloseAndReadBack(wb1);
            cell = wb2.GetSheet("bug58778").GetRow(0).GetCell(0);
            Assert.AreEqual(5.25, cell.NumericCellValue, 0);

            style = cell.CellStyle;
            Assert.AreEqual(poundFmt, style.GetDataFormatString());
            Assert.AreEqual(poundFmtIdx, style.DataFormat);

            // manually check the file to make sure the cell is rendered as "<poundsymbol>5"
            // Verified with LibreOffice 4.2.8.2 on 2015-12-28
            wb2.Close();
            wb1.Close();
        }

    }

}



