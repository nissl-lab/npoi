/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
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


using NPOI.Util;
using NPOI.XSSF.EventUserModel;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XSSF.EventUserModel
{
    using NPOI.OpenXml4Net.OPC;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.Binary;
    using NPOI.XSSF.UserModel;

    [TestFixture]
    public class TestXSSFBReader
    {
        private static POIDataSamples _ssTests = POIDataSamples.GetSpreadSheetInstance();

        [Test]
        public void TestBasic()
        {

            List<String> sheetTexts = GetSheets("testVarious.xlsb");

            Assert.AreEqual(1, sheetTexts.Count);
            String xsxml = sheetTexts[0];
            POITestCase.AssertContains(xsxml, "This is a string");
            POITestCase.AssertContains(xsxml, "<td ref=\"B2\">13</td>");
            POITestCase.AssertContains(xsxml, "<td ref=\"B3\">13.12112313</td>");
            POITestCase.AssertContains(xsxml, "<td ref=\"B4\">$   3.03</td>");
            POITestCase.AssertContains(xsxml, "<td ref=\"B5\">20%</td>");
            POITestCase.AssertContains(xsxml, "<td ref=\"B6\">13.12</td>");
            POITestCase.AssertContains(xsxml, "<td ref=\"B7\">1.23457E+14</td>");
            POITestCase.AssertContains(xsxml, "<td ref=\"B8\">1.23457E+15</td>");

            POITestCase.AssertContains(xsxml, "46/1963");//custom format 1
            POITestCase.AssertContains(xsxml, "3/128");//custom format 2

            POITestCase.AssertContains(xsxml, "<tr num=\"7>\n" +
                    "\t<td ref=\"A8\">longer int</td>\n" +
                    "\t<td ref=\"B8\">1.23457E+15</td>\n" +
                    "\t<td ref=\"C8\"><span type=\"comment\" author=\"Allison, Timothy B.\">Allison, Timothy B.:\n" +
                    "test comment2</span></td>\n" +
                    "</tr num=\"7>");

            POITestCase.AssertContains(xsxml, "<tr num=\"34>\n" +
                    "\t<td ref=\"B35\">comment6<span type=\"comment\" author=\"Allison, Timothy B.\">Allison, Timothy B.:\n" +
                    "comment6 actually in cell</span></td>\n" +
                    "</tr num=\"34>");

            POITestCase.AssertContains(xsxml, "<tr num=\"64>\n" +
                    "\t<td ref=\"I65\"><span type=\"comment\" author=\"Allison, Timothy B.\">Allison, Timothy B.:\n" +
                    "comment7 end of file</span></td>\n" +
                    "</tr num=\"64>");

            POITestCase.AssertContains(xsxml, "<tr num=\"65>\n" +
                    "\t<td ref=\"I66\"><span type=\"comment\" author=\"Allison, Timothy B.\">Allison, Timothy B.:\n" +
                    "comment8 end of file</span></td>\n" +
                    "</tr num=\"65>");

            POITestCase.AssertContains(xsxml,
                    "<header tagName=\"header\">OddLeftHeader OddCenterHeader OddRightHeader</header>");
            POITestCase.AssertContains(xsxml,
                    "<footer tagName=\"footer\">OddLeftFooter OddCenterFooter OddRightFooter</footer>");
            POITestCase.AssertContains(xsxml,
                    "<header tagName=\"evenHeader\">EvenLeftHeader EvenCenterHeader EvenRightHeader\n</header>");
            POITestCase.AssertContains(xsxml,
                    "<footer tagName=\"evenFooter\">EvenLeftFooter EvenCenterFooter EvenRightFooter</footer>");
            POITestCase.AssertContains(xsxml,
                    "<header tagName=\"firstHeader\">FirstPageLeftHeader FirstPageCenterHeader FirstPageRightHeader</header>");
            POITestCase.AssertContains(xsxml,
                    "<footer tagName=\"firstFooter\">FirstPageLeftFooter FirstPageCenterFooter FirstPageRightFooter</footer>");

        }

        [Test]
        public void TestComments()
        {

            List<String> sheetTexts = GetSheets("comments.xlsb");
            String xsxml = sheetTexts[0];
            POITestCase.AssertContains(xsxml,
                    "<tr num=\"0>\n" +
                            "\t<td ref=\"A1\"><span type=\"comment\" author=\"Sven Nissel\">comment top row1 (index0)</span></td>\n" +
                            "\t<td ref=\"B1\">row1</td>\n" +
                            "</tr num=\"0>");
            POITestCase.AssertContains(xsxml,
                    "<tr num=\"1>\n" +
                            "\t<td ref=\"A2\"><span type=\"comment\" author=\"Allison, Timothy B.\">Allison, Timothy B.:\n" +
                            "comment row2 (index1)</span></td>\n" +
                            "</tr num=\"1>");
            POITestCase.AssertContains(xsxml, "<tr num=\"2>\n" +
                    "\t<td ref=\"A3\">row3<span type=\"comment\" author=\"Sven Nissel\">comment top row3 (index2)</span></td>\n" +
                    "\t<td ref=\"B3\">row3</td>\n");

            POITestCase.AssertContains(xsxml, "<tr num=\"3>\n" +
                    "\t<td ref=\"A4\"><span type=\"comment\" author=\"Sven Nissel\">comment top row4 (index3)</span></td>\n" +
                    "\t<td ref=\"B4\">row4</td>\n" +
                    "</tr num=\"3></sheet>");

        }

        [Test]
        public void TestAbsPath()
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("testVarious.xlsb"));
            XSSFBReader r = new XSSFBReader(pkg);
            Assert.AreEqual("C:\\Users\\tallison\\Desktop\\working\\xlsb\\", r.GetAbsPathMetadata());
        }

        private List<String> GetSheets(String testFileName)
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream(testFileName));
            List<String> sheetTexts = new List<String>();
            XSSFBReader r = new XSSFBReader(pkg);

            //        Assert.IsNotNull(r.WorkbookData);
            //      Assert.IsNotNull(r.SharedStringsData);
            Assert.IsNotNull(r.GetXSSFBStylesTable());
            XSSFBSharedStringsTable sst = new XSSFBSharedStringsTable(pkg);
            XSSFBStylesTable xssfbStylesTable = r.GetXSSFBStylesTable();
            XSSFBReader.SheetIterator itr = (XSSFBReader.SheetIterator) r.GetSheetsData();

            while(itr.MoveNext())
            {
                Stream is1 = itr.Current;
                String name = itr.SheetName;
                TestSheetHandler testSheetHandler = new TestSheetHandler();
                testSheetHandler.StartSheet(name);
                XSSFBSheetHandler sheetHandler = new XSSFBSheetHandler(is1,
                    xssfbStylesTable,
                    itr.GetXSSFBSheetComments(),
                    sst, testSheetHandler,
                    new DataFormatter(),
                    false);
                sheetHandler.Parse();
                testSheetHandler.EndSheet();
                sheetTexts.Add(testSheetHandler.ToString());
            }
            return sheetTexts;

        }

        [Test]
        public void TestDate()
        {

            List<String> sheets = GetSheets("date.xlsb");
            Assert.AreEqual(1, sheets.Count);
            POITestCase.AssertContains(sheets[0], "1/12/13");
        }


        private class TestSheetHandler : XSSFSheetXMLHandler.ISheetContentsHandler
        {
            private  StringBuilder sb = new StringBuilder();

            public void StartSheet(String sheetName)
            {
                sb.Append("<sheet name=\"").Append(sheetName).Append(">");
            }

            public void EndSheet()
            {
                sb.Append("</sheet>");
            }
            public void StartRow(int rowNum)
            {
                sb.Append("\n<tr num=\"").Append(rowNum).Append(">");
            }
            public void EndRow(int rowNum)
            {
                sb.Append("\n</tr num=\"").Append(rowNum).Append(">");
            }
            public void Cell(String cellReference, String formattedValue, XSSFComment comment)
            {
                formattedValue = (formattedValue == null) ? "" : formattedValue;
                if(comment == null)
                {
                    sb.Append("\n\t<td ref=\"").Append(cellReference).Append("\">").Append(formattedValue).Append("</td>");
                }
                else
                {
                    sb.Append("\n\t<td ref=\"").Append(cellReference).Append("\">")
                            .Append(formattedValue)
                            .Append("<span type=\"comment\" author=\"")
                            .Append(comment.Author).Append("\">")
                            .Append(comment.String.ToString().Trim()).Append("</span>")
                            .Append("</td>");
                }
            }
            public void HeaderFooter(String text, bool IsHeader, String tagName)
            {
                if(IsHeader)
                {
                    sb.Append("<header tagName=\"" + tagName + "\">" + text + "</header>");
                }
                else
                {
                    sb.Append("<footer tagName=\"" + tagName + "\">" + text + "</footer>");

                }
            }
            public override String ToString()
            {
                return sb.ToString();
            }
        }
    }
}

