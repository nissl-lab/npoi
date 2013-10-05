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

using NPOI.SS.UserModel;
using System.Collections.Generic;
using System;
using NUnit.Framework;
using NPOI.HSSF.UserModel;
namespace TestCases.SS.UserModel
{

    /**
     * Tests of implementation of {@link DataFormat}
     *
     */
    public class BaseTestDataFormat
    {

        protected ITestDataProvider _testDataProvider;
        public BaseTestDataFormat()
            : this(TestCases.HSSF.HSSFITestDataProvider.Instance)
        { }
        /**
         * @param testDataProvider an object that provides test data in HSSF / XSSF specific way
         */
        protected BaseTestDataFormat(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }
        [Test]
        public void TestBuiltinFormats()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();

            IDataFormat df = wb.CreateDataFormat();

            List<String> formats = HSSFDataFormat.GetBuiltinFormats();
            for (int idx = 0; idx < formats.Count; idx++)
            {
                String fmt = formats[idx];
                Assert.AreEqual(idx, df.GetFormat(fmt));
            }

            //default format for new cells is General
            ISheet sheet = wb.CreateSheet();
            ICell cell = sheet.CreateRow(0).CreateCell(0);
            Assert.AreEqual(0, cell.CellStyle.DataFormat);
            Assert.AreEqual("General", cell.CellStyle.GetDataFormatString());

            //create a custom data format
            String customFmt = "#0.00 AM/PM";
            //check it is not in built-in formats
            Assert.AreEqual(-1, HSSFDataFormat.GetBuiltinFormat(customFmt));
            int customIdx = df.GetFormat(customFmt);
            //The first user-defined format starts at 164.
            Assert.IsTrue(customIdx >= HSSFDataFormat.FIRST_USER_DEFINED_FORMAT_INDEX);
            //read and verify the string representation
            Assert.AreEqual(customFmt, df.GetFormat((short)customIdx));
        }

        /**
         * [Bug 49928] formatCellValue returns incorrect value for \u00a3 formatted cells
         */
        public virtual void Test49928()
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook("49928.xls");
            doTest49928Core(wb);
        }
        protected String poundFmt = "\"\u00a3\"#,##0;[Red]\\-\"\u00a3\"#,##0";
        public void doTest49928Core(IWorkbook wb)
        {
            DataFormatter df = new DataFormatter();

            ISheet sheet = wb.GetSheetAt(0);
            ICell cell = sheet.GetRow(0).GetCell(0);
            ICellStyle style = cell.CellStyle;

            String poundFmt = "\"\u00a3\"#,##0;[Red]\\-\"\u00a3\"#,##0";
            // not expected normally, id of a custom format should be greater 
            // than BuiltinFormats.FIRST_USER_DEFINED_FORMAT_INDEX
            short poundFmtIdx = 6;

            Assert.AreEqual(poundFmt, style.GetDataFormatString());
            Assert.AreEqual(poundFmtIdx, style.DataFormat);
            Assert.AreEqual("\u00a31", df.FormatCellValue(cell));


            IDataFormat dataFormat = wb.CreateDataFormat();
            Assert.AreEqual(poundFmtIdx, dataFormat.GetFormat(poundFmt));
            Assert.AreEqual(poundFmt, dataFormat.GetFormat(poundFmtIdx));
        }
    }
}




