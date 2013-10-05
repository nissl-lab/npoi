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

using TestCases.SS;
using NUnit.Framework;
using NPOI.SS.UserModel;
namespace TestCases.SS.UserModel
{

    /**
     * Test diffrent types of Excel hyperlinks
     *
     * @author Yegor Kozlov
     */
    public class BaseTestHyperlink
    {

        protected ITestDataProvider _testDataProvider;
        public BaseTestHyperlink()
            : this(TestCases.HSSF.HSSFITestDataProvider.Instance)
        {}

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
            IWorkbook wb = _testDataProvider.CreateWorkbook();
            ICreationHelper CreateHelper = wb.GetCreationHelper();

            ICell cell;
            IHyperlink link;
            ISheet sheet = wb.CreateSheet("Hyperlinks");

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
            ISheet sheet2 = wb.CreateSheet("Target Sheet");
            sheet2.CreateRow(0).CreateCell((short)0).SetCellValue("Target Cell");

            cell = sheet.CreateRow(3).CreateCell((short)0);
            cell.SetCellValue("Worksheet Link");
            link = CreateHelper.CreateHyperlink(HyperlinkType.Document);
            link.Address = ("'Target Sheet'!A1");
            cell.Hyperlink = (link);

            wb = _testDataProvider.WriteOutAndReadBack(wb);

            sheet = wb.GetSheetAt(0);
            link = sheet.GetRow(0).GetCell(0).Hyperlink;

            Assert.AreEqual("http://poi.apache.org/", link.Address);
            link = sheet.GetRow(1).GetCell(0).Hyperlink;
            Assert.AreEqual("hyperinks-beta4-dump.txt", link.Address);
            link = sheet.GetRow(2).GetCell(0).Hyperlink;
            Assert.AreEqual("mailto:poi@apache.org?subject=Hyperlinks", link.Address);
            link = sheet.GetRow(3).GetCell(0).Hyperlink;
            Assert.AreEqual("'Target Sheet'!A1", link.Address);
        }
    }

}


