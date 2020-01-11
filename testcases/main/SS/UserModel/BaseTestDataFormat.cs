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

        public void AssertNotBuiltInFormat(String customFmt)
        {
            //check it is not in built-in formats
            Assert.AreEqual(-1, BuiltinFormats.GetBuiltinFormat(customFmt));
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
            AssertNotBuiltInFormat(customFmt);
            int customIdx = df.GetFormat(customFmt);
            //The first user-defined format starts at 164.
            Assert.IsTrue(customIdx >= HSSFDataFormat.FIRST_USER_DEFINED_FORMAT_INDEX);
            //read and verify the string representation
            Assert.AreEqual(customFmt, df.GetFormat((short)customIdx));

            wb.Close();
        }

        /**
         * [Bug 49928] formatCellValue returns incorrect value for \u00a3 formatted cells
         */
        public virtual void Test49928()
        {
            IWorkbook wb = _testDataProvider.OpenSampleWorkbook("49928.xls");
            DoTest49928Core(wb);
        }
        protected String poundFmt = "\"\u00a3\"#,##0;[Red]\\-\"\u00a3\"#,##0";
        public void DoTest49928Core(IWorkbook wb)
        {
            DataFormatter df = new DataFormatter();

            ISheet sheet = wb.GetSheetAt(0);
            ICell cell = sheet.GetRow(0).GetCell(0);
            ICellStyle style = cell.CellStyle;

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

        [Test]
        public void TestReadbackFormat()
        {
            ReadbackFormat("built-in format", "0.00");
            ReadbackFormat("overridden built-in format", poundFmt);

            String customFormat = "#0.00 AM/PM";
            AssertNotBuiltInFormat(customFormat);
            ReadbackFormat("custom format", customFormat);
        }

        private void ReadbackFormat(String msg, String fmt)
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                IDataFormat dataFormat = wb.CreateDataFormat();
                short fmtIdx = dataFormat.GetFormat(fmt);
                String readbackFmt = dataFormat.GetFormat(fmtIdx);
                Assert.AreEqual(fmt, readbackFmt, msg);
            }
            finally
            {
                wb.Close();
            }
        }

        public void DoTest58532Core(IWorkbook wb)
        {
            ISheet s = wb.GetSheetAt(0);
            DataFormatter fmt = new DataFormatter();
            IFormulaEvaluator eval = wb.GetCreationHelper().CreateFormulaEvaluator();

            // Column A is the raw values
            // Column B is the ##/#K/#M values
            // Column C is strings of what they should look like
            // Column D is the #.##/#.#K/#.#M values
            // Column E is strings of what they should look like

            String formatKMWhole = "[>999999]#,,\"M\";[>999]#,\"K\";#";
            String formatKM3dp = "[>999999]#.000,,\"M\";[>999]#.000,\"K\";#.000";

            // Check the formats are as expected
            IRow headers = s.GetRow(0);
            Assert.IsNotNull(headers);
            Assert.AreEqual(formatKMWhole, headers.GetCell(1).StringCellValue);
            Assert.AreEqual(formatKM3dp, headers.GetCell(3).StringCellValue);

            IRow r2 = s.GetRow(1);
            Assert.IsNotNull(r2);
            Assert.AreEqual(formatKMWhole, r2.GetCell(1).CellStyle.GetDataFormatString());
            Assert.AreEqual(formatKM3dp, r2.GetCell(3).CellStyle.GetDataFormatString());

            // For all of the contents rows, check that DataFormatter is able
            //  to format the cells to the same value as the one next to it
            for (int rn = 1; rn < s.LastRowNum; rn++)
            {
                IRow r = s.GetRow(rn);
                if (r == null) break;

                double value = r.GetCell(0).NumericCellValue;

                String expWhole = r.GetCell(2).StringCellValue;
                String exp3dp = r.GetCell(4).StringCellValue;

                Assert.AreEqual(expWhole, fmt.FormatCellValue(r.GetCell(1), eval), "Wrong formatting of " + value + " for row " + rn);
                Assert.AreEqual(exp3dp, fmt.FormatCellValue(r.GetCell(3), eval), "Wrong formatting of " + value + " for row " + rn);
            }
        }

        /**
         * Localised accountancy formats
         */
         [Test]
        public void Test58536()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            DataFormatter formatter = new DataFormatter();
            IDataFormat fmt = wb.CreateDataFormat();
            ISheet sheet = wb.CreateSheet();
            IRow r = sheet.CreateRow(0);

            char pound = '\u00A3';
            String formatUK = "_-[$" + pound + "-809]* #,##0_-;\\-[$" + pound + "-809]* #,##0_-;_-[$" + pound + "-809]* \"-\"??_-;_-@_-";

            ICellStyle cs = wb.CreateCellStyle();
            cs.DataFormat = (fmt.GetFormat(formatUK));

            ICell pve = r.CreateCell(0);
            pve.SetCellValue(12345);
            pve.CellStyle = (cs);

            ICell nve = r.CreateCell(1);
            nve.SetCellValue(-12345);
            nve.CellStyle = (cs);

            ICell zero = r.CreateCell(2);
            zero.SetCellValue(0);
            zero.CellStyle = (cs);

            Assert.AreEqual(pound + "   12,345", formatter.FormatCellValue(pve));
            Assert.AreEqual("-" + pound + "   12,345", formatter.FormatCellValue(nve));
            // TODO Fix this to not have an extra 0 at the end
            //assertEquals(pound+"   -  ", formatter.formatCellValue(zero)); 

            wb.Close();
        }


        /**
         * Using a single quote (') instead of a comma (,) as
         *  a number separator, eg 1000 -> 1'000 
         */
        [Test]
        public void Test55265()
        {

            IWorkbook wb = _testDataProvider.CreateWorkbook();
            try
            {
                DataFormatter formatter = new DataFormatter();
                IDataFormat fmt = wb.CreateDataFormat();
                ISheet sheet = wb.CreateSheet();
                IRow r = sheet.CreateRow(0);

                ICellStyle cs = wb.CreateCellStyle();
                cs.DataFormat = (fmt.GetFormat("#'##0"));

                ICell zero = r.CreateCell(0);
                zero.SetCellValue(0);
                zero.CellStyle = (cs);

                ICell sml = r.CreateCell(1);
                sml.SetCellValue(12);
                sml.CellStyle = (cs);

                ICell med = r.CreateCell(2);
                med.SetCellValue(1234);
                med.CellStyle = (cs);

                ICell lge = r.CreateCell(3);
                lge.SetCellValue(12345678);
                lge.CellStyle = (cs);

                Assert.AreEqual("0", formatter.FormatCellValue(zero));
                Assert.AreEqual("12", formatter.FormatCellValue(sml));
                Assert.AreEqual("1'234", formatter.FormatCellValue(med));
                Assert.AreEqual("12'345'678", formatter.FormatCellValue(lge));
            }
            finally { 
                wb.Close();
            }
        }
    }
}




