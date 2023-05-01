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
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF;

namespace TestCases.XSSF.Model
{
    [TestFixture]
    public class TestStylesTable
    {
        private String testFile = "Formatting.xlsx";
        private static String customDataFormat = "YYYY-mm-dd";

        [SetUp]
        public static void assumeCustomDataFormatIsNotBuiltIn()
        {
            Assert.AreEqual(-1, BuiltinFormats.GetBuiltinFormat(customDataFormat));
        }

        [Test]
        public void TestCreateNew()
        {
            StylesTable st = new StylesTable();

            // Check defaults
            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(1, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(0, st.NumDataFormats);
        }
        [Test]
        public void TestCreateSaveLoad()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            StylesTable st = wb.GetStylesSource();

            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(1, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(0, st.NumDataFormats);

            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb)).GetStylesSource();

            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(1, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(0, st.NumDataFormats);

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        [Test]
        public void TestLoadExisting()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook(testFile);
            Assert.IsNotNull(workbook.GetStylesSource());

            StylesTable st = workbook.GetStylesSource();

            doTestExisting(st);

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
        }
        [Test]
        public void TestLoadSaveLoad()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook(testFile);
            Assert.IsNotNull(workbook.GetStylesSource());

            StylesTable st = workbook.GetStylesSource();
            doTestExisting(st);

            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook)).GetStylesSource();
            doTestExisting(st);
        }
        public void doTestExisting(StylesTable st)
        {
            // Check contents
            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(11, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(8, st.NumDataFormats);

            Assert.AreEqual(2, st.GetFonts().Count);
            Assert.AreEqual(2, st.GetFills().Count);
            Assert.AreEqual(1, st.GetBorders().Count);

            Assert.AreEqual("yyyy/mm/dd", st.GetNumberFormatAt((short)165));
            Assert.AreEqual("yy/mm/dd", st.GetNumberFormatAt((short)167));

            Assert.IsNotNull(st.GetStyleAt(0));
            Assert.IsNotNull(st.GetStyleAt(1));
            Assert.IsNotNull(st.GetStyleAt(2));

            Assert.AreEqual(0, st.GetStyleAt(0).DataFormat);
            Assert.AreEqual(14, st.GetStyleAt(1).DataFormat);
            Assert.AreEqual(0, st.GetStyleAt(2).DataFormat);
            Assert.AreEqual(165, st.GetStyleAt(3).DataFormat);

            Assert.AreEqual("yyyy/mm/dd", st.GetStyleAt(3).GetDataFormatString());
        }
        [Test]
        public void TestPopulateNew()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            StylesTable st = wb.GetStylesSource();

            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(1, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(0, st.NumDataFormats);

            int nf1 = st.PutNumberFormat("yyyy-mm-dd");
            int nf2 = st.PutNumberFormat("yyyy-mm-DD");
            Assert.AreEqual(nf1, st.PutNumberFormat("yyyy-mm-dd"));

            st.PutStyle(new XSSFCellStyle(st));

            // Save and re-load
            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb)).GetStylesSource();

            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(2, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(2, st.NumDataFormats);

            Assert.AreEqual("yyyy-mm-dd", st.GetNumberFormatAt((short)nf1));
            Assert.AreEqual(nf1, st.PutNumberFormat("yyyy-mm-dd"));
            Assert.AreEqual(nf2, st.PutNumberFormat("yyyy-mm-DD"));

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(wb));
        }
        [Test]
        public void TestPopulateExisting()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook(testFile);
            Assert.IsNotNull(workbook.GetStylesSource());

            StylesTable st = workbook.GetStylesSource();
            Assert.AreEqual(11, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(8, st.NumDataFormats);

            int nf1 = st.PutNumberFormat("YYYY-mm-dd");
            int nf2 = st.PutNumberFormat("YYYY-mm-DD");
            Assert.AreEqual(nf1, st.PutNumberFormat("YYYY-mm-dd"));

            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook)).GetStylesSource();

            Assert.AreEqual(11, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(10, st.NumDataFormats);

            Assert.AreEqual("YYYY-mm-dd", st.GetNumberFormatAt((short)nf1));
            Assert.AreEqual(nf1, st.PutNumberFormat("YYYY-mm-dd"));
            Assert.AreEqual(nf2, st.PutNumberFormat("YYYY-mm-DD"));

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
        }

        [Test]
        public void ExceedNumberFormatLimit()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                StylesTable styles = wb.GetStylesSource();
                for (int i = 0; i < styles.MaxNumberOfDataFormats; i++)
                {
                    wb.GetStylesSource().PutNumberFormat("\"test" + i + " \"0");
                }
                try
                {
                    wb.GetStylesSource().PutNumberFormat("\"anotherformat \"0");
                }
                catch (InvalidOperationException e)
                {
                    if (e.Message.StartsWith("The maximum number of Data Formats was exceeded."))
                    {
                        //expected
                    }
                    else
                    {
                        throw e;
                    }
                }
            }
            finally
            {
                wb.Close();
            }
        }

        private static void assertNotContainsKey<K, V>(SortedDictionary<K, V> map, K key)
        {
            Assert.IsFalse(map.ContainsKey(key));
        }
        private static void assertNotContainsValue<K, V>(SortedDictionary<K, V> map, V value)
        {
            Assert.IsFalse(map.ContainsValue(value));
        }

        [Test]
        public void removeNumberFormat()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            try
            {
                String fmt = customDataFormat;
                short fmtIdx = (short)wb.GetStylesSource().PutNumberFormat(fmt);

                ICell cell = wb.CreateSheet("test").CreateRow(0).CreateCell(0);
                cell.SetCellValue(5.25);
                ICellStyle style = wb.CreateCellStyle();
                style.DataFormat = fmtIdx;
                cell.CellStyle = style;

                Assert.AreEqual(fmt, cell.CellStyle.GetDataFormatString());
                Assert.AreEqual(fmt, wb.GetStylesSource().GetNumberFormatAt(fmtIdx));

                // remove the number format from the workbook
                wb.GetStylesSource().RemoveNumberFormat(fmt);

                // number format in CellStyles should be restored to default number format
                short defaultFmtIdx = 0;
                String defaultFmt = BuiltinFormats.GetBuiltinFormat(0);
                Assert.AreEqual(defaultFmtIdx, style.DataFormat);
                Assert.AreEqual(defaultFmt, style.GetDataFormatString());

                // The custom number format should be entirely removed from the workbook
                SortedDictionary<short, String> numberFormats = wb.GetStylesSource().GetNumberFormats() as SortedDictionary<short, String>;
                assertNotContainsKey(numberFormats, fmtIdx);
                assertNotContainsValue(numberFormats, fmt);

                // The default style shouldn't be added back to the styles source because it's built-in
                Assert.AreEqual(0, wb.GetStylesSource().NumDataFormats);

                cell = null;
                style = null;
                numberFormats = null;
                wb = XSSFTestDataSamples.WriteOutCloseAndReadBack(wb);

                cell = wb.GetSheet("test").GetRow(0).GetCell(0);
                style = cell.CellStyle;

                // number format in CellStyles should be restored to default number format
                Assert.AreEqual(defaultFmtIdx, style.DataFormat);
                Assert.AreEqual(defaultFmt, style.GetDataFormatString());

                // The custom number format should be entirely removed from the workbook
                numberFormats = wb.GetStylesSource().GetNumberFormats() as SortedDictionary<short, String>;
                assertNotContainsKey(numberFormats, fmtIdx);
                assertNotContainsValue(numberFormats, fmt);

                // The default style shouldn't be added back to the styles source because it's built-in
                Assert.AreEqual(0, wb.GetStylesSource().NumDataFormats);

            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void maxNumberOfDataFormats()
        {
            XSSFWorkbook wb = new XSSFWorkbook();

            try
            {
                StylesTable styles = wb.GetStylesSource();

                // Check default limit
                int n = styles.MaxNumberOfDataFormats;
                // https://support.office.com/en-us/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
                Assert.IsTrue(200 <= n);
                Assert.IsTrue(n <= 250);

                // Check upper limit
                n = int.MaxValue;
                styles.MaxNumberOfDataFormats = (n);
                Assert.AreEqual(n, styles.MaxNumberOfDataFormats);

                // Check negative (illegal) limits
                try
                {
                    styles.MaxNumberOfDataFormats = (-1);
                    Assert.Fail("Expected to get an IllegalArgumentException(\"Maximum Number of Data Formats must be greater than or equal to 0\")");
                }
                catch (ArgumentException e)
                {
                    if (e.Message.StartsWith("Maximum Number of Data Formats must be greater than or equal to 0"))
                    {
                        // expected
                    }
                    else
                    {
                        throw e;
                    }
                }
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void addDataFormatsBeyondUpperLimit()
        {
            XSSFWorkbook wb = new XSSFWorkbook();

            try
            {
                StylesTable styles = wb.GetStylesSource();
                styles.MaxNumberOfDataFormats = (0);

                // Try adding a format beyond the upper limit
                try
                {
                    styles.PutNumberFormat("\"test \"0");
                    Assert.Fail("Expected to raise InvalidOperationException");
                }
                catch (InvalidOperationException e)
                {
                    if (e.Message.StartsWith("The maximum number of Data Formats was exceeded."))
                    {
                        // expected
                    }
                    else
                    {
                        throw e;
                    }
                }
            }
            finally
            {
                wb.Close();
            }
        }

        [Test]
        public void decreaseUpperLimitBelowCurrentNumDataFormats()
        {
            XSSFWorkbook wb = new XSSFWorkbook();

            try
            {
                StylesTable styles = wb.GetStylesSource();
                styles.PutNumberFormat(customDataFormat);

                // Try decreasing the upper limit below the current number of formats
                try
                {
                    styles.MaxNumberOfDataFormats = (0);
                    Assert.Fail("Expected to raise InvalidOperationException");
                }
                catch (InvalidOperationException e)
                {
                    if (e.Message.StartsWith("Cannot set the maximum number of data formats less than the current quantity."))
                    {
                        // expected
                    }
                    else
                    {
                        throw e;
                    }
                }
            }
            finally
            {
                wb.Close();
            }
        }

    }
}
