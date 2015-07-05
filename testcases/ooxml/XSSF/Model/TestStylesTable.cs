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
namespace NPOI.XSSF.Model
{
    [TestFixture]
    public class TestStylesTable
    {
        private String testFile = "Formatting.xlsx";
        [Test]
        public void TestCreateNew()
        {
            StylesTable st = new StylesTable();

            // Check defaults
            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(1, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(0, st.NumberFormatSize);
        }
        [Test]
        public void TestCreateSaveLoad()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            StylesTable st = wb.GetStylesSource();

            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(1, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(0, st.NumberFormatSize);

            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb)).GetStylesSource();

            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(1, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(0, st.NumberFormatSize);

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
            Assert.AreEqual(8, st.NumberFormatSize);

            Assert.AreEqual(2, st.GetFonts().Count);
            Assert.AreEqual(2, st.GetFills().Count);
            Assert.AreEqual(1, st.GetBorders().Count);

            Assert.AreEqual("yyyy/mm/dd", st.GetNumberFormatAt(165));
            Assert.AreEqual("yy/mm/dd", st.GetNumberFormatAt(167));

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
            Assert.AreEqual(0, st.NumberFormatSize);

            int nf1 = st.PutNumberFormat("yyyy-mm-dd");
            int nf2 = st.PutNumberFormat("yyyy-mm-DD");
            Assert.AreEqual(nf1, st.PutNumberFormat("yyyy-mm-dd"));

            st.PutStyle(new XSSFCellStyle(st));

            // Save and re-load
            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb)).GetStylesSource();

            Assert.IsNotNull(st.GetCTStylesheet());
            Assert.AreEqual(2, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(2, st.NumberFormatSize);

            Assert.AreEqual("yyyy-mm-dd", st.GetNumberFormatAt(nf1));
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
            Assert.AreEqual(8, st.NumberFormatSize);

            int nf1 = st.PutNumberFormat("YYYY-mm-dd");
            int nf2 = st.PutNumberFormat("YYYY-mm-DD");
            Assert.AreEqual(nf1, st.PutNumberFormat("YYYY-mm-dd"));

            st = ((XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook)).GetStylesSource();

            Assert.AreEqual(11, st.XfsSize);
            Assert.AreEqual(1, st.StyleXfsSize);
            Assert.AreEqual(10, st.NumberFormatSize);

            Assert.AreEqual("YYYY-mm-dd", st.GetNumberFormatAt(nf1));
            Assert.AreEqual(nf1, st.PutNumberFormat("YYYY-mm-dd"));
            Assert.AreEqual(nf2, st.PutNumberFormat("YYYY-mm-DD"));

            Assert.IsNotNull(XSSFTestDataSamples.WriteOutAndReadBack(workbook));
        }
    }
}
