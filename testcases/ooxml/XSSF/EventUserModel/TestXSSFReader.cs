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


using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XSSF.EventUserModel
{
    using NPOI;
    using NPOI.OpenXml4Net.OPC;
    using NPOI.Util;
    using NPOI.XSSF;
    using NPOI.XSSF.Model;
    using NPOI.XSSF.UserModel;
    using NPOI.XSSF.EventUserModel;
    using NUnit.Framework;


    /// <summary>
    /// Tests for <see cref="XSSFReader"/>
    /// </summary>
    [TestFixture]
    public sealed class TestXSSFReader
    {
        private static POIDataSamples _ssTests = POIDataSamples.GetSpreadSheetInstance();

        [Test]
        public void TestGetBits()
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("SampleSS.xlsx"));

            XSSFReader r = new XSSFReader(pkg);

            Assert.IsNotNull(r.WorkbookData);
            Assert.IsNotNull(r.SharedStringsData);
            Assert.IsNotNull(r.StylesData);

            Assert.IsNotNull(r.SharedStringsTable);
            Assert.IsNotNull(r.StylesTable);
        }

        [Test]
        public void TestStyles()
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("SampleSS.xlsx"));

            XSSFReader r = new XSSFReader(pkg);

            Assert.AreEqual(3, r.StylesTable.Fonts.Count);
            Assert.AreEqual(0, r.StylesTable.NumDataFormats);

            // The Styles Table should have the themes associated with it too
            Assert.IsNotNull(r.StylesTable.Theme);

            // Check we Get valid data for the two
            Assert.IsNotNull(r.StylesData);
            Assert.IsNotNull(r.ThemesData);
        }

        [Test]
        public void TestStrings()
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("SampleSS.xlsx"));

            XSSFReader r = new XSSFReader(pkg);

            Assert.AreEqual(11, r.SharedStringsTable.Items.Count);
            Assert.AreEqual("Test spreadsheet", new XSSFRichTextString(r.SharedStringsTable.GetEntryAt(0)).ToString());
        }

        [Test]
        public void TestSheets()
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("SampleSS.xlsx"));

            XSSFReader r = new XSSFReader(pkg);
            byte[] data = new byte[4096];

            // By r:id
            Assert.IsNotNull(r.GetSheet("rId2"));
            int read = IOUtils.ReadFully(r.GetSheet("rId2"), data);
            Assert.AreEqual(974, read);

            // All
            IEnumerator<Stream> it = r.GetSheetsData();

            int count = 0;
            while(it.MoveNext())
            {
                count++;
                Stream inp = it.Current;
                Assert.IsNotNull(inp);
                read = IOUtils.ReadFully(inp, data);
                inp.Close();

                Assert.IsTrue(read > 400);
                Assert.IsTrue(read < 1500);
            }
            Assert.AreEqual(3, count);
        }

        /// <summary>
        /// Check that the sheet iterator returns sheets in the logical order
        /// (as they are defined in the workbook.xml)
        /// </summary>
        [Test]
        public void TestOrderOfSheets()
        {

            OPCPackage pkg = OPCPackage.Open(_ssTests.OpenResourceAsStream("reordered_sheets.xlsx"));

            XSSFReader r = new XSSFReader(pkg);

            String[] sheetNames = {"Sheet4", "Sheet2", "Sheet3", "Sheet1"};
            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator)r.GetSheetsData();

            int count = 0;
            while(it.MoveNext())
            {
                Stream inp = it.Current;
                Assert.IsNotNull(inp);
                inp.Close();

                Assert.AreEqual(sheetNames[count], it.SheetName);
                count++;
            }
            Assert.AreEqual(4, count);
        }
        [Test]
        public void TestComments()
        {

            OPCPackage pkg =  XSSFTestDataSamples.OpenSamplePackage("comments.xlsx");
            XSSFReader r = new XSSFReader(pkg);
            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator)r.GetSheetsData();

            int count = 0;
            while(it.MoveNext())
            {
                count++;
                Stream inp = it.Current;
                inp.Close();

                if(count == 1)
                {
                    Assert.IsNotNull(it.SheetComments);
                    CommentsTable ct = it.SheetComments;
                    Assert.AreEqual(1, ct.NumberOfAuthors);
                    Assert.AreEqual(3, ct.NumberOfComments);
                }
                else
                {
                    Assert.IsNull(it.SheetComments);
                }
            }
            Assert.AreEqual(3, count);
        }

        /// <summary>
        /// Iterating over a workbook with chart sheets in it, using the
        ///  XSSFReader method
        /// </summary>
        /// <exception cref="Exception">Exception</exception>
        [Test]
        public void Test50119()
        {

            OPCPackage pkg =  XSSFTestDataSamples.OpenSamplePackage("WithChartSheet.xlsx");
            XSSFReader r = new XSSFReader(pkg);
            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator)r.GetSheetsData();

            while(it.MoveNext())
            {
                Stream stream = it.Current;
                stream.Close();
            }
        }

        /// <summary>
        /// Test text extraction from text box using GetShapes()
        /// </summary>
        /// <exception cref="Exception">Exception</exception>
        [Test]
        public void TestShapes()
        {

            OPCPackage pkg = XSSFTestDataSamples.OpenSamplePackage("WithTextBox.xlsx");
            XSSFReader r = new XSSFReader(pkg);
            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) r.GetSheetsData();

            String text = GetShapesString(it);
            StringAssert.Contains("Line 1", text);
            StringAssert.Contains("Line 2", text);
            StringAssert.Contains("Line 3", text);
        }

        private String GetShapesString(XSSFReader.SheetIterator it)
        {
            StringBuilder sb = new StringBuilder();
            while(it.MoveNext())
            {
                var _ = it.Current;
                List<XSSFShape> shapes = it.Shapes;
                if(shapes != null)
                {
                    foreach(XSSFShape shape in shapes)
                    {
                        if(shape is XSSFSimpleShape)
                        {
                            String t = ((XSSFSimpleShape) shape).Text;
                            sb.Append(t).Append('\n');
                        }
                    }
                }
            }
            return sb.ToString();
        }
        [Test]
        public void TestBug57914()
        {

            OPCPackage pkg = XSSFTestDataSamples.OpenSamplePackage("57914.xlsx");
            XSSFReader r;

            // for now expect this to Assert.Fail, when we fix 57699, this one should Assert.Fail so we know we should adjust
            // this test as well
            try
            {
                r = new XSSFReader(pkg);
                Assert.Fail("This will Assert.Fail until bug 57699 is fixed");
            }
            catch(POIXMLException e)
            {
                StringAssert.Contains("57699", e.Message);
                return;
            }

            XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) r.GetSheetsData();

            String text = GetShapesString(it);
            StringAssert.Contains("Line 1", text);
            StringAssert.Contains("Line 2", text);
            StringAssert.Contains("Line 3", text);
        }

        /// <summary>
        /// NPE from XSSFReader$SheetIterator.<init> on XLSX files generated by
        ///  the openpyxl library
        /// </summary>
        [Test]
        public void Test58747()
        {

            OPCPackage pkg =  XSSFTestDataSamples.OpenSamplePackage("58747.xlsx");
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
            Assert.IsNotNull(strings);
            XSSFReader reader = new XSSFReader(pkg);
            StylesTable styles = reader.StylesTable;
            Assert.IsNotNull(styles);

            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.GetSheetsData();
            Assert.AreEqual(true, iter.MoveNext());
            var _ = iter.Current;

            Assert.AreEqual(false, iter.MoveNext());
            Assert.AreEqual("Orders", iter.SheetName);

            pkg.Close();
        }

        /// <summary>
        /// NPE when sheet has no relationship id in the workbook
        /// 60825
        /// </summary>
        [Test]
        public void TestSheetWithNoRelationshipId()
        {

            OPCPackage pkg =  XSSFTestDataSamples.OpenSamplePackage("60825.xlsx");
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
            Assert.IsNotNull(strings);
            XSSFReader reader = new XSSFReader(pkg);
            StylesTable styles = reader.StylesTable;
            Assert.IsNotNull(styles);

            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.GetSheetsData();
            iter.MoveNext();
            Assert.IsNotNull(iter.Current);
            Assert.IsFalse(iter.MoveNext());

            pkg.Close();
        }

        /// <summary>
        /// <para>
        /// bug 61304: Call to XSSFReader.SheetsData returns duplicate sheets.
        /// </para>
        /// <para>
        /// The problem seems to be caused only by those xlsx files which have a specific
        /// order of the attributes inside the &lt;sheet&gt; tag of workbook.xml
        /// </para>
        /// <para>
        /// Example (which causes the problems):
        /// &lt;sheet name="Sheet6" r:id="rId6" sheetId="4"/&gt;
        /// </para>
        /// <para>
        /// While this one works correctly:
        /// &lt;sheet name="Sheet6" sheetId="4" r:id="rId6"/&gt;
        /// </para>
        /// </summary>
        [Test]
        public void Test61034()
        {
            OPCPackage pkg = XSSFTestDataSamples.OpenSamplePackage("61034.xlsx");
            XSSFReader reader = new XSSFReader(pkg);
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.GetSheetsData();
            ISet<String> seen = new HashSet<String>();
            while(iter.MoveNext())
            {
                Stream stream = iter.Current;
                String sheetName = iter.SheetName;
                CollectionAssert.DoesNotContain(seen, sheetName);
                seen.Add(sheetName);
                stream.Close();
            }
            pkg.Close();
        }
    }
}

