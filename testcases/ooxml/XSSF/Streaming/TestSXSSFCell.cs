/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for Additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

namespace TestCases.XSSF.Streaming
{
    using NPOI.OpenXmlFormats.Spreadsheet;
    using NPOI.SS.UserModel;
    using NPOI.XSSF;
    using NPOI.XSSF.UserModel;
    using NUnit.Framework;
    using System;
    using TestCases.SS.UserModel;

    /**
     * Tests various functionality having to do with {@link SXSSFCell}.  For instance support for
     * particular datatypes, etc.
     */
    [TestFixture]
    public class TestSXSSFCell : BaseTestXCell
    {

        public TestSXSSFCell()
            : base(SXSSFITestDataProvider.instance)
        {

        }

        [TearDown]
        public static void TearDown() {
            //SXSSFITestDataProvider.instance.Cleanup();
        }

        [Test]
        public void TestPreserveSpaces() {
            String[] samplesWithSpaces = {
                " POI",
                "POI ",
                " POI ",
                "\nPOI",
                "\n\nPOI \n",
        };
            foreach (String str in samplesWithSpaces) {
                IWorkbook swb = _testDataProvider.CreateWorkbook();
                ICell sCell = swb.CreateSheet().CreateRow(0).CreateCell(0);
                sCell.SetCellValue(str);
                Assert.AreEqual(sCell.StringCellValue, str);

                // read back as XSSF and check that xml:spaces="preserve" is Set
                XSSFWorkbook xwb = (XSSFWorkbook)_testDataProvider.WriteOutAndReadBack(swb);
                XSSFCell xCell = xwb.GetSheetAt(0).GetRow(0).GetCell(0) as XSSFCell;

                CT_Rst is1 = xCell.GetCTCell().@is;
                //XmlCursor c = is1.NewCursor();
                //c.ToNextToken();
                //String t = c.GetAttributeText(new QName("http://www.w3.org/XML/1998/namespace", "space"));
                //c.Dispose();


                //write is1 to xml stream writer ,get the xml text and parse it and get space attr.
                //Assert.AreEqual("preserve", t, "expected xml:spaces=\"preserve\" \"" + str + "\"");
                xwb.Close();
                swb.Close();
            }
        }
    }

}