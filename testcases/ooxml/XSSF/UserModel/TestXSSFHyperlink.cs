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
using NPOI.OpenXml4Net.OPC;
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
        public void TestCreate()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.CreateSheet() as XSSFSheet;
            XSSFRow row = sheet.CreateRow(0) as XSSFRow;
            XSSFCreationHelper CreateHelper = workbook.GetCreationHelper() as XSSFCreationHelper;

            String[] urls = {
                "http://apache.org/",
                "www.apache.org",
                "/temp",
                "file:///c:/temp",
                "http://apache.org/default.php?s=isTramsformed&submit=Search&la=*&li=*"};
            for (int i = 0; i < urls.Length; i++)
            {
                String s = urls[i];
                XSSFHyperlink link = CreateHelper.CreateHyperlink(HyperlinkType.Url) as XSSFHyperlink;
                link.Address=(s);

                XSSFCell cell = row.CreateCell(i) as XSSFCell;
                cell.Hyperlink=(link);
            }
            workbook = XSSFTestDataSamples.WriteOutAndReadBack(workbook) as XSSFWorkbook;
            sheet = workbook.GetSheetAt(0) as XSSFSheet;
            PackageRelationshipCollection rels = sheet.GetPackagePart().Relationships;
            Assert.AreEqual(urls.Length, rels.Size);
            for (int i = 0; i < rels.Size; i++)
            {
                PackageRelationship rel = rels.GetRelationship(i);
                if (rel.TargetUri.IsAbsoluteUri&&rel.TargetUri.IsFile)
                    Assert.AreEqual(urls[i].Replace("file:///","").Replace("/","\\"),rel.TargetUri.LocalPath);
                else
                    // there should be a relationship for each URL
                    Assert.AreEqual(urls[i], rel.TargetUri.ToString());
            }

            // Bugzilla 53041: Hyperlink relations are duplicated when saving XSSF file
            workbook = XSSFTestDataSamples.WriteOutAndReadBack(workbook) as XSSFWorkbook;
            sheet = workbook.GetSheetAt(0) as XSSFSheet;
            rels = sheet.GetPackagePart().Relationships;
            Assert.AreEqual(urls.Length, rels.Size);
            for (int i = 0; i < rels.Size; i++)
            {
                PackageRelationship rel = rels.GetRelationship(i);
                if (rel.TargetUri.IsAbsoluteUri && rel.TargetUri.IsFile)
                    Assert.AreEqual(urls[i].Replace("file:///", "").Replace("/", "\\"), rel.TargetUri.LocalPath);
                else
                    // there should be a relationship for each URL
                    Assert.AreEqual(urls[i], rel.TargetUri.ToString());
            }
        }
        [Test]
        public void TestInvalidURLs()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFCreationHelper CreateHelper = workbook.GetCreationHelper() as XSSFCreationHelper;

            String[] invalidURLs = {
                "http:\\apache.org",
                "www.apache .org",
                "c:\\temp",
                "\\poi"};
            foreach (String s in invalidURLs)
            {
                try
                {
                    CreateHelper.CreateHyperlink(HyperlinkType.Url).Address = (s);
                    Assert.Fail("expected ArgumentException: " + s);
                }
                catch (ArgumentException)
                {
                }
            }
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

            IHyperlink hyperlink = CreateHelper.CreateHyperlink(HyperlinkType.Url);
            hyperlink.Address = ("http://poi.apache.org/spreadsheet/");
            hyperlink.Label = "POI SS Link";
            r17c.Hyperlink=(hyperlink);

            Assert.AreEqual(5, sheet.NumHyperlinks);
            doTestHyperlinkContents(sheet);

            Assert.AreEqual(HyperlinkType.Url,
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

            Assert.AreEqual(HyperlinkType.Url,
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
            Assert.AreEqual(HyperlinkType.Url,
                    sheet.GetRow(3).GetCell(2).Hyperlink.Type);
            Assert.AreEqual(null,
                    sheet.GetRow(3).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("http://poi.apache.org/",
                    sheet.GetRow(3).GetCell(2).Hyperlink.Address);

            // Next is an internal doc link
            Assert.AreEqual(HyperlinkType.Document,
                    sheet.GetRow(14).GetCell(2).Hyperlink.Type);
            Assert.AreEqual("Internal hyperlink to A2",
                    sheet.GetRow(14).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("Sheet1!A2",
                    sheet.GetRow(14).GetCell(2).Hyperlink.Address);

            // Next is a file
            Assert.AreEqual(HyperlinkType.File,
                    sheet.GetRow(15).GetCell(2).Hyperlink.Type);
            Assert.AreEqual(null,
                    sheet.GetRow(15).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("WithVariousData.xlsx",
                    sheet.GetRow(15).GetCell(2).Hyperlink.Address);

            // Last is a mailto
            Assert.AreEqual(HyperlinkType.Email,
                    sheet.GetRow(16).GetCell(2).Hyperlink.Type);
            Assert.AreEqual(null,
                    sheet.GetRow(16).GetCell(2).Hyperlink.Label);
            Assert.AreEqual("mailto:dev@poi.apache.org?subject=XSSF Hyperlinks",
                    sheet.GetRow(16).GetCell(2).Hyperlink.Address);
        }
        [Test]
        public void Test52716()
        {
            XSSFWorkbook wb1 = XSSFTestDataSamples.OpenSampleWorkbook("52716.xlsx");
            XSSFSheet sh1 = wb1.GetSheetAt(0) as XSSFSheet;

            XSSFWorkbook wb2 = XSSFTestDataSamples.WriteOutAndReadBack(wb1) as XSSFWorkbook;
            XSSFSheet sh2 = wb2.GetSheetAt(0) as XSSFSheet;

            Assert.AreEqual(sh1.NumberOfComments, sh2.NumberOfComments);
            XSSFHyperlink l1 = sh1.GetHyperlink(0, 1);
            Assert.AreEqual(HyperlinkType.Document, l1.Type);
            Assert.AreEqual("B1", l1.GetCellRef());
            Assert.AreEqual("Sort on Titel", l1.Tooltip);

            XSSFHyperlink l2 = sh2.GetHyperlink(0, 1);
            Assert.AreEqual(l1.Tooltip, l2.Tooltip);
            Assert.AreEqual(HyperlinkType.Document, l2.Type);
            Assert.AreEqual("B1", l2.GetCellRef());
        }
        [Test]
        public void Test53734()
        {
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53734.xlsx");
            XSSFHyperlink link = wb.GetSheetAt(0).GetRow(0).GetCell(0).Hyperlink as XSSFHyperlink;
            Assert.AreEqual("javascript:///", link.Address);

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            link = wb.GetSheetAt(0).GetRow(0).GetCell(0).Hyperlink as XSSFHyperlink;
            Assert.AreEqual("javascript:///", link.Address);
        }
        [Test]
        public void Test53282()
        {
            //since limitation in .NET Uri class, it's impossible to accept uri like mailto:nobody@nowhere.uk%C2%A0
            XSSFWorkbook wb = XSSFTestDataSamples.OpenSampleWorkbook("53282.xlsx");
            XSSFHyperlink link = wb.GetSheetAt(0).GetRow(0).GetCell(14).Hyperlink as XSSFHyperlink;
            Assert.AreEqual("mailto:nobody@nowhere.uk%C2%A0", link.Address);

            wb = XSSFTestDataSamples.WriteOutAndReadBack(wb) as XSSFWorkbook;
            link = wb.GetSheetAt(0).GetRow(0).GetCell(14).Hyperlink as XSSFHyperlink;
            Assert.AreEqual("mailto:nobody@nowhere.uk%C2%A0", link.Address);
        }

    }


}

