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

namespace TestCases.HSSF.UserModel
{
    using System;
    using System.IO;
    using NUnit.Framework;
    using TestCases.HSSF;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using TestCases.SS.UserModel;
    /**
     * Tests HSSFHyperlink.
     *
     * @author  Yegor Kozlov
     */
    [TestFixture]
    public class TestHSSFHyperlink : BaseTestHyperlink
    {
        public TestHSSFHyperlink()
            : base(HSSFITestDataProvider.Instance)
        {
            
        }
        /**
         * Test that we can read hyperlinks.
         */
        [Test]
        public void TestRead()
        {

            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("HyperlinksOnManySheets.xls");

            NPOI.SS.UserModel.ISheet sheet;
            ICell cell;
            IHyperlink link;

            sheet = wb.GetSheet("WebLinks");
            cell = sheet.GetRow(4).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("POI", link.Label);
            Assert.AreEqual("POI", cell.RichStringCellValue.String);
            Assert.AreEqual("http://poi.apache.org/", link.Address);

            cell = sheet.GetRow(8).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("HSSF", link.Label);
            Assert.AreEqual("HSSF", cell.RichStringCellValue.String);
            Assert.AreEqual("http://poi.apache.org/hssf/", link.Address);

            sheet = wb.GetSheet("Emails");
            cell = sheet.GetRow(4).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("dev", link.Label);
            Assert.AreEqual("dev", cell.RichStringCellValue.String);
            Assert.AreEqual("mailto:dev@poi.apache.org", link.Address);

            sheet = wb.GetSheet("Internal");
            cell = sheet.GetRow(4).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("Link To First Sheet", link.Label);
            Assert.AreEqual("Link To First Sheet", cell.RichStringCellValue.String);
            Assert.AreEqual("WebLinks!A1", link.TextMark);
            Assert.AreEqual("WebLinks!A1", link.Address);
        }
        [Test]
        public void TestModify()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("HyperlinksOnManySheets.xls");

            NPOI.SS.UserModel.ISheet sheet;
            ICell cell;
            IHyperlink link;

            sheet = wb.GetSheet("WebLinks");
            cell = sheet.GetRow(4).GetCell(0);
            link = cell.Hyperlink;
            //modify the link
            link.Address = ("www.apache.org");

            //serialize and read again
            MemoryStream out1 = new MemoryStream();
            wb.Write(out1);

            wb = new HSSFWorkbook(new MemoryStream(out1.ToArray()));
            sheet = wb.GetSheet("WebLinks");
            cell = sheet.GetRow(4).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("www.apache.org", link.Address);

        }
        /**
         * HSSF-specific ways of creating links to a place in workbook.
         * You can set the target  in two ways:
         *  link.SetTextMark("'Target Sheet-1'!A1"); //HSSF-specific
         *  or
         *  link.SetAddress("'Target Sheet-1'!A1"); //common between XSSF and HSSF
         */
        [Test]
        public void TestCreateDocumentLink()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            //link to a place in this workbook
            IHyperlink link;
            ICell cell;
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Hyperlinks");

            //create a target sheet and cell
            NPOI.SS.UserModel.ISheet sheet2 = wb.CreateSheet("Target Sheet");
            sheet2.CreateRow(0).CreateCell(0).SetCellValue("Target Cell");

            //cell A1 has a link to 'Target Sheet-1'!A1
            cell = sheet.CreateRow(0).CreateCell(0);
            cell.SetCellValue("Worksheet Link");
            link = new HSSFHyperlink(HyperlinkType.Document);
            link.TextMark=("'Target Sheet'!A1");
            cell.Hyperlink=(link);

            //cell B1 has a link to cell A1 on the same sheet
            cell = sheet.CreateRow(1).CreateCell(0);
            cell.SetCellValue("Worksheet Link");
            link = new HSSFHyperlink(HyperlinkType.Document);
            link.Address=("'Hyperlinks'!A1");
            cell.Hyperlink=(link);

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sheet = wb.GetSheet("Hyperlinks");

            cell = sheet.GetRow(0).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("'Target Sheet'!A1", link.TextMark);
            Assert.AreEqual("'Target Sheet'!A1", link.Address);

            cell = sheet.GetRow(1).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("'Hyperlinks'!A1", link.TextMark);
            Assert.AreEqual("'Hyperlinks'!A1", link.Address);
        }
                /**
         * Test that NPOI.SS.UserModel.Sheet#shiftRows moves hyperlinks,
         * see bugs #46445 and #29957
         */
        [Test]
        public void TestShiftRows()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("46445.xls");


            NPOI.SS.UserModel.ISheet sheet = wb.GetSheetAt(0);

            //verify existing hyperlink in A3
            ICell cell1 = sheet.GetRow(2).GetCell(0);
            IHyperlink link1 = cell1.Hyperlink;
            Assert.IsNotNull(link1);
            Assert.AreEqual(2, link1.FirstRow);
            Assert.AreEqual(2, link1.LastRow);

            //assign a hyperlink to A4
            HSSFHyperlink link2 = new HSSFHyperlink(HyperlinkType.Document);
            link2.Address=("Sheet2!A2");
            ICell cell2 = sheet.GetRow(3).GetCell(0);
            cell2.Hyperlink=(link2);
            Assert.AreEqual(3, link2.FirstRow);
            Assert.AreEqual(3, link2.LastRow);

            //move the 3rd row two rows down
            sheet.ShiftRows(sheet.FirstRowNum, sheet.LastRowNum, 2);

            //cells A3 and A4 don't contain hyperlinks anymore
            Assert.IsNull(sheet.GetRow(2).GetCell(0).Hyperlink);
            Assert.IsNull(sheet.GetRow(3).GetCell(0).Hyperlink);

            //the first hypelink now belongs to A5
            IHyperlink link1_shifted = sheet.GetRow(2 + 2).GetCell(0).Hyperlink;
            Assert.IsNotNull(link1_shifted);
            Assert.AreEqual(4, link1_shifted.FirstRow);
            Assert.AreEqual(4, link1_shifted.LastRow);

            //the second hypelink now belongs to A6
            IHyperlink link2_shifted = sheet.GetRow(3 + 2).GetCell(0).Hyperlink;
            Assert.IsNotNull(link2_shifted);
            Assert.AreEqual(5, link2_shifted.FirstRow);
            Assert.AreEqual(5, link2_shifted.LastRow);
        }
        [Test]
        public void TestCreate()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            ICell cell;
            NPOI.SS.UserModel.ISheet sheet = wb.CreateSheet("Hyperlinks");

            //URL
            cell = sheet.CreateRow(0).CreateCell(0);
            cell.SetCellValue("URL Link");
            IHyperlink link = new HSSFHyperlink(HyperlinkType.Url);
            link.Address = ("http://poi.apache.org/");
            cell.Hyperlink = link;

            //link to a file in the current directory
            cell = sheet.CreateRow(1).CreateCell(0);
            cell.SetCellValue("File Link");
            link = new HSSFHyperlink(HyperlinkType.File);
            link.Address = ("link1.xls");
            cell.Hyperlink = link;

            //e-mail link
            cell = sheet.CreateRow(2).CreateCell(0);
            cell.SetCellValue("Email Link");
            link = new HSSFHyperlink(HyperlinkType.Email);
            //note, if subject contains white spaces, make sure they are url-encoded
            link.Address = ("mailto:poi@apache.org?subject=Hyperlinks");
            cell.Hyperlink = link;

            //link to a place in this workbook

            //Create a target sheet and cell
            NPOI.SS.UserModel.ISheet sheet2 = wb.CreateSheet("Target Sheet");
            sheet2.CreateRow(0).CreateCell(0).SetCellValue("Target Cell");

            cell = sheet.CreateRow(3).CreateCell(0);
            cell.SetCellValue("Worksheet Link");
            link = new HSSFHyperlink(HyperlinkType.Document);
            link.TextMark = ("'Target Sheet'!A1");
            cell.Hyperlink = link;

            //serialize and read again
            MemoryStream out1 = new MemoryStream();
            wb.Write(out1);

            wb = new HSSFWorkbook(new MemoryStream(out1.ToArray()));
            sheet = wb.GetSheet("Hyperlinks");
            cell = sheet.GetRow(0).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("http://poi.apache.org/", link.Address);

            cell = sheet.GetRow(1).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("link1.xls", link.Address);

            cell = sheet.GetRow(2).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("mailto:poi@apache.org?subject=Hyperlinks", link.Address);

            cell = sheet.GetRow(3).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("'Target Sheet'!A1", link.TextMark);
        }
        [Test]
        public void TestCloneSheet()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("HyperlinksOnManySheets.xls");

            ICell cell;
            IHyperlink link;

            NPOI.SS.UserModel.ISheet sheet = wb.CloneSheet(0);

            cell = sheet.GetRow(4).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("http://poi.apache.org/", link.Address);

            cell = sheet.GetRow(8).GetCell(0);
            link = cell.Hyperlink;
            Assert.IsNotNull(link);
            Assert.AreEqual("http://poi.apache.org/hssf/", link.Address);
        }

        public override IHyperlink CopyHyperlink(IHyperlink link)
        {
            return new HSSFHyperlink(link);
        }

        /*
        [Test]
        public void testCopyXSSFHyperlink() {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFCreationHelper helper = wb.CreationHelper;
            XSSFHyperlink xlink = helper.createHyperlink(Hyperlink.LINK_URL);
            xlink.setAddress("http://poi.apache.org/");
            xlink.setCellReference("D3");
            xlink.setTooltip("tooltip");
            HSSFHyperlink hlink = new HSSFHyperlink(xlink);

            Assert.AreEqual("http://poi.apache.org/", hlink.Address);
            Assert.AreEqual("D3", new CellReference(hlink.FirstRow, hlink.FirstColumn).formatAsString());
            // Are HSSFHyperlink.label and XSSFHyperlink.tooltip the same? If so, perhaps one of these needs renamed for a consistent Hyperlink interface
            // Assert.AreEqual("tooltip", hlink.Label);

            wb.close();
        }*/
    }
}