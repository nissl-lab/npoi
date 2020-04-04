/*
* Licensed to the Apache Software Foundation (ASF) under one or more
* contributor license agreements.  See the NOTICE file distributed with
* this work for Additional information regarding copyright ownership.
* The ASF licenses this file to You under the Apache License, Version 2.0
* (the "License"); you may not use this file except in compliance with
* the License.  You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*/
namespace TestCases.HSSF.UserModel
{
    using NUnit.Framework;

    using TestCases.HSSF;
    using TestCases.SS.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.HSSF.UserModel;
    using System;

    /**
     * Test <c>HSSFPicture</c>.
     *
     * @author Yegor Kozlov (yegor at apache.org)
     */
    [TestFixture]
    public class TestHSSFDataFormat : BaseTestDataFormat
    {
        public TestHSSFDataFormat()
            : base(HSSFITestDataProvider.Instance)
        {

        }
        /**
        * [Bug 49928] formatCellValue returns incorrect value for \u00a3 formatted cells
        */
        [Test]
        public new void Test49928()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("49928.xls");
            DoTest49928Core(wb);

            // an attempt to register an existing format returns its index
            int poundFmtIdx = wb.GetSheetAt(0).GetRow(0).GetCell(0).CellStyle.DataFormat;
            Assert.AreEqual(poundFmtIdx, wb.CreateDataFormat().GetFormat(poundFmt));

            // now create a custom format with Pound (\u00a3)
            IDataFormat dataFormat = wb.CreateDataFormat();
            short customFmtIdx = dataFormat.GetFormat("\u00a3##.00[Yellow]");
            Assert.IsTrue(customFmtIdx >= BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX);
            Assert.AreEqual("\u00a3##.00[Yellow]", dataFormat.GetFormat(customFmtIdx));

            wb.Close();
        }
        /**
         * [Bug 58532] Handle formats that go numnum, numK, numM etc 
         */
        [Test]
        public void test58532()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("FormatKM.xls");
            DoTest58532Core(wb);
            wb.Close();
        }
        /**
         * Bug 51378: GetDataFormatString method call crashes when Reading the test file
         */
        [Test]
        public void Test51378()
        {
            IWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("12561-1.xls");
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                ISheet sheet = wb.GetSheetAt(i);
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row)
                    {
                        ICellStyle style = cell.CellStyle;

                        String fmt = style.GetDataFormatString();
                        if (fmt == null)
                        {
                            //_logger.Log(POILogger.WARN, cell + ": " + fmt);
                        }
                    }
                }
            }
            wb.Close();
        }
    }
}