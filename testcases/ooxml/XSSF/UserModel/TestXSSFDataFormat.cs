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
namespace NPOI.XSSF.UserModel
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
            doTest49928Core(wb);

            // an attempt to register an existing format returns its index
            int poundFmtIdx = wb.GetSheetAt(0).GetRow(0).GetCell(0).CellStyle.DataFormat;
            Assert.AreEqual(poundFmtIdx, wb.GetStylesSource().PutNumberFormat(poundFmt));

            // now create a custom format with Pound (\u00a3)
            IDataFormat dataFormat = wb.CreateDataFormat();
            short customFmtIdx = dataFormat.GetFormat("\u00a3##.00[Yellow]");
            Assert.IsTrue(customFmtIdx > BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX);
            Assert.AreEqual("\u00a3##.00[Yellow]", dataFormat.GetFormat(customFmtIdx));
        }
    }

}



