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
using System;
using NUnit.Framework;
using NPOI.SS.UserModel;
namespace NPOI.XSSF.UserModel
{
    [TestFixture]
    public class TestXSSFHyperlink : BaseTestHyperlink
    {
        public TestXSSFHyperlink()
            : base(XSSFITestDataProvider.instance)
        {

        }

        [SetUp]
        public void SetUp()
        {
            // Use system out logger
            Environment.SetEnvironmentVariable(
                    "NPOI.util.POILogger",
                    "NPOI.util.SystemOutLogger"
            );
        }
        [Test]
        public void TestLoadExisting()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("WithMoreVariousData.xlsx");
            Assert.AreEqual(3, workbook.NumberOfSheets);

            XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0);

            // Check the hyperlinks
            Assert.AreEqual(4, sheet.NumHyperlinks);
            doTestHyperlinkContents(sheet);
        }
        [Test]
        public void TestLoadSave()
        {
            XSSFWorkbook workbook = XSSFTestDataSamples.OpenSampleWorkbook("WithMoreVariousData.xlsx");
            ICreationHelper CreateHelper = workbook.GetCreationHelper();
            Assert.AreEqual(3, workbook.NumberOfSheets);
            XSSFSheet sheet = (XSSFSheet)workbook.GetSheetAt(0);

            // Check hyperlinks
            Assert.AreEqual(4, sheet.NumHyperlinks);
            doTestHyperlinkContents(sheet);


            // Write out, and check

            // Load up again, check all links still there
            XSSFWorkbook wb2 = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(workbook);
            Assert.AreEqual(3, wb2.NumberOfSheets);
            Assert.IsNotNull(wb2.GetSheetAt(0));
            Assert.IsNotNull(wb2.GetSheetAt(1));
            Assert.IsNotNull(wb2.GetSheetAt(2));

            sheet = (XSSFSheet)wb2.GetSheetAt(0);


            // Check hyperlinks again
            Assert.AreEqual(4, sheet.NumHyperlinks);
            doTestHyperlinkContents(sheet);


            // Add one more, and re-check
            IRow r17 = sheet.CreateRow(17);
            ICell r17c = r17.CreateCell(2);

            IHyperlink hyperlink = CreateHelper.CreateHyperlink(HyperlinkType.URL);
            hyperlink.Address = ("http://poi.apache.org/spreadsheet/");
            hyperlink.Label = "POI SS Link";
            r17c.Hyperlink=(hyperlink);

            Assert.AreEqual(5, sheet.NumHyperlinks);
            doTestHyperlinkContents(sheet);

            Assert.AreEqual(HyperlinkType.URL,
                    sheet.GetRow(17).GetCell(2).Hyperlink.Type);
            Assert.AreEqual("POI SS Link",
                    sheet.GetRow(17).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("http://poi.apache.org/spreadsheet/",
                    sheet.GetRow(17).GetCell(2).Hyperlink.Address);


            // Save and re-load once more

            XSSFWorkbook wb3 = (XSSFWorkbook)XSSFTestDataSamples.WriteOutAndReadBack(wb2);
            Assert.AreEqual(3, wb3.NumberOfSheets);
            Assert.IsNotNull(wb3.GetSheetAt(0));
            Assert.IsNotNull(wb3.GetSheetAt(1));
            Assert.IsNotNull(wb3.GetSheetAt(2));

            sheet = (XSSFSheet)wb3.GetSheetAt(0);

            Assert.AreEqual(5, sheet.NumHyperlinks);
            doTestHyperlinkContents(sheet);

            Assert.AreEqual(HyperlinkType.URL,
                    sheet.GetRow(17).GetCell(2).Hyperlink.Type);
            Assert.AreEqual("POI SS Link",
                    sheet.GetRow(17).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("http://poi.apache.org/spreadsheet/",
                    sheet.GetRow(17).GetCell(2).Hyperlink.Address);
        }

        /**
         * Only for WithMoreVariousData.xlsx !
         */
        private static void doTestHyperlinkContents(XSSFSheet sheet)
        {
            Assert.IsNotNull(sheet.GetRow(3).GetCell(2).Hyperlink);
            Assert.IsNotNull(sheet.GetRow(14).GetCell(2).Hyperlink);
            Assert.IsNotNull(sheet.GetRow(15).GetCell(2).Hyperlink);
            Assert.IsNotNull(sheet.GetRow(16).GetCell(2).Hyperlink);

            // First is a link to poi
            Assert.AreEqual(HyperlinkType.URL,
                    sheet.GetRow(3).GetCell(2).Hyperlink.Type);
            Assert.AreEqual(null,
                    sheet.GetRow(3).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("http://poi.apache.org/",
                    sheet.GetRow(3).GetCell(2).Hyperlink.Address);

            // Next is an internal doc link
            Assert.AreEqual(HyperlinkType.DOCUMENT,
                    sheet.GetRow(14).GetCell(2).Hyperlink.Type);
            Assert.AreEqual("Internal hyperlink to A2",
                    sheet.GetRow(14).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("Sheet1!A2",
                    sheet.GetRow(14).GetCell(2).Hyperlink.Address);

            // Next is a file
            Assert.AreEqual(HyperlinkType.FILE,
                    sheet.GetRow(15).GetCell(2).Hyperlink.Type);
            Assert.AreEqual(null,
                    sheet.GetRow(15).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("WithVariousData.xlsx",
                    sheet.GetRow(15).GetCell(2).Hyperlink.Address);

            // Last is a mailto
            Assert.AreEqual(HyperlinkType.EMAIL,
                    sheet.GetRow(16).GetCell(2).Hyperlink.Type);
            Assert.AreEqual(null,
                    sheet.GetRow(16).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("mailto:dev@poi.apache.org?subject=XSSF Hyperlinks",
                    sheet.GetRow(16).GetCell(2).Hyperlink.Address);
        }
    }


}

