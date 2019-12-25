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

using NUnit.Framework;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System;

namespace TestCases.SS.UserModel
{

    /**
     * Test diffrent types of Excel hyperlinks
     *
     * @author Yegor Kozlov
     */
    public abstract class BaseTestHyperlink
    {

        protected ITestDataProvider _testDataProvider;

        /**
         * @param testDataProvider an object that provides test data in HSSF / XSSF specific way
         */
        protected BaseTestHyperlink(ITestDataProvider testDataProvider)
        {
            _testDataProvider = testDataProvider;
        }
        [Test]
        public void TestBasicTypes()
        {
            IWorkbook wb1 = _testDataProvider.CreateWorkbook();
            ICreationHelper CreateHelper = wb1.GetCreationHelper();

            ICell cell;
            IHyperlink link;
            ISheet sheet = wb1.CreateSheet("Hyperlinks");

            //URL
            cell = sheet.CreateRow(0).CreateCell((short)0);
            cell.SetCellValue("URL Link");
            link = CreateHelper.CreateHyperlink(HyperlinkType.Url);
            link.Address = ("http://poi.apache.org/");
            cell.Hyperlink = (link);

            //link to a file in the current directory
            cell = sheet.CreateRow(1).CreateCell((short)0);
            cell.SetCellValue("File Link");
            link = CreateHelper.CreateHyperlink(HyperlinkType.File);
            link.Address = ("hyperinks-beta4-dump.txt");
            cell.Hyperlink = (link);

            //e-mail link
            cell = sheet.CreateRow(2).CreateCell((short)0);
            cell.SetCellValue("Email Link");
            link = CreateHelper.CreateHyperlink(HyperlinkType.Email);
            //note, if subject Contains white spaces, make sure they are url-encoded
            link.Address = ("mailto:poi@apache.org?subject=Hyperlinks");
            cell.Hyperlink = (link);

            //link to a place in this workbook

            //create a target sheet and cell
            ISheet sheet2 = wb1.CreateSheet("Target Sheet");
            sheet2.CreateRow(0).CreateCell((short)0).SetCellValue("Target Cell");

            cell = sheet.CreateRow(3).CreateCell((short)0);
            cell.SetCellValue("Worksheet Link");
            link = CreateHelper.CreateHyperlink(HyperlinkType.Document);
            link.Address = ("'Target Sheet'!A1");
            cell.Hyperlink = (link);

            IWorkbook wb2 = _testDataProvider.WriteOutAndReadBack(wb1);

            sheet = wb2.GetSheetAt(0);
            link = sheet.GetRow(0).GetCell(0).Hyperlink;

            Assert.AreEqual("http://poi.apache.org/", link.Address);
            link = sheet.GetRow(1).GetCell(0).Hyperlink;
            Assert.AreEqual("hyperinks-beta4-dump.txt", link.Address);
            link = sheet.GetRow(2).GetCell(0).Hyperlink;
            Assert.AreEqual("mailto:poi@apache.org?subject=Hyperlinks", link.Address);
            link = sheet.GetRow(3).GetCell(0).Hyperlink;
            Assert.AreEqual("'Target Sheet'!A1", link.Address);

            wb2.Close();
        }

        // copy a hyperlink via the copy constructor
        [Test]
        public void TestCopyHyperlink()
        {
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICreationHelper createHelper = wb.GetCreationHelper();
            ISheet sheet = wb.CreateSheet("Hyperlinks");
            IRow row = sheet.CreateRow(0);
            ICell cell1, cell2;
            IHyperlink link1, link2;
            //URL
            cell1 = row.CreateCell(0);
            cell2 = row.CreateCell(1);
            cell1.SetCellValue("URL Link");
            link1 = createHelper.CreateHyperlink(HyperlinkType.Url);
            link1.Address = ("http://poi.apache.org/");
            cell1.Hyperlink = (link1);

            link2 = CopyHyperlink(link1);

            // Change address (type is not changeable)
            link2.Address = ("http://apache.org/");
            cell2.Hyperlink = (link2);

            // Make sure hyperlinks were deep-copied, and modifying one does not modify the other. 
            Assert.AreNotSame(link1, link2);
            Assert.AreNotEqual(link1, link2);
            Assert.AreEqual("http://poi.apache.org/", link1.Address);
            Assert.AreEqual("http://apache.org/", link2.Address);
            Assert.AreEqual(link1, cell1.Hyperlink);
            Assert.AreEqual(link2, cell2.Hyperlink);

            // Make sure both hyperlinks were added to the sheet
            List<IHyperlink> actualHyperlinks = sheet.GetHyperlinkList();
            Assert.AreEqual(2, actualHyperlinks.Count);
            Assert.AreEqual(link1, actualHyperlinks[0]);
            Assert.AreEqual(link2, actualHyperlinks[1]);

            wb.Close();
        }

        public abstract IHyperlink CopyHyperlink(IHyperlink link);
    }

}


