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
namespace NPOI.XWPF.UserModel
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NPOI.XWPF;
    using NPOI.OpenXmlFormats.Wordprocessing;
    using System.Collections.Generic;


    /**
     * Tests for XWPF Run
     */
    [TestClass]
    public class TestXWPFTable
    {

        protected void SetUp()
        {
            /*
              XWPFDocument doc = new XWPFDocument();
              p = doc.CreateParagraph();

              this.ctRun = CTR.Factory.NewInstance();
           */
        }

        [TestMethod]
        public void TestConstructor()
        {
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl ctTable = new CT_Tbl();
            XWPFTable xtab = new XWPFTable(ctTable, doc);
            Assert.IsNotNull(xtab);
            Assert.AreEqual(1, ctTable.SizeOfTrArray());
            Assert.AreEqual(1, ctTable.GetTrArray(0).SizeOfTcArray());
            Assert.IsNotNull(ctTable.GetTrArray(0).GetTcArray(0).GetPArray(0));

            ctTable = new CT_Tbl();
            xtab = new XWPFTable(ctTable, doc, 3, 2);
            Assert.IsNotNull(xtab);
            Assert.AreEqual(3, ctTable.SizeOfTrArray());
            Assert.AreEqual(2, ctTable.GetTrArray(0).SizeOfTcArray());
            Assert.IsNotNull(ctTable.GetTrArray(0).GetTcArray(0).GetPArray(0));
        }


        [TestMethod]
        public void TestGetText()
        {
            XWPFDocument doc = new XWPFDocument();
            CT_Tbl table = new CT_Tbl();
            CT_Row row = table.AddNewTr();
            CT_Tc cell = row.AddNewTc();
            CT_P paragraph = cell.AddNewP();
            CT_R run = paragraph.AddNewR();
            CT_Text text = run.AddNewT();
            text.Value = ("finally I can Write!");

            XWPFTable xtab = new XWPFTable(table, doc);
            Assert.AreEqual("finally I can Write!\n", xtab.GetText());
        }


        [TestMethod]
        public void TestCreateRow()
        {
            XWPFDocument doc = new XWPFDocument();

            CT_Tbl table = new CT_Tbl();
            CT_Row r1 = table.AddNewTr();
            r1.AddNewTc().AddNewP();
            r1.AddNewTc().AddNewP();
            CT_Row r2 = table.AddNewTr();
            r2.AddNewTc().AddNewP();
            r2.AddNewTc().AddNewP();
            CT_Row r3 = table.AddNewTr();
            r3.AddNewTc().AddNewP();
            r3.AddNewTc().AddNewP();

            XWPFTable xtab = new XWPFTable(table, doc);
            Assert.AreEqual(3, xtab.GetNumberOfRows());
            Assert.IsNotNull(xtab.GetRow(2));

            //add a new row
            xtab.CreateRow();
            Assert.AreEqual(4, xtab.GetNumberOfRows());

            //check number of cols
            Assert.AreEqual(2, table.GetTrArray(0).SizeOfTcArray());

            //check creation of first row
            xtab = new XWPFTable(new CT_Tbl(), doc);
            Assert.AreEqual(1, xtab.GetCTTbl().GetTrArray(0).SizeOfTcArray());
        }

        [TestMethod]
        public void TestSetWidth()
        {
            XWPFDocument doc = new XWPFDocument();

            CT_Tbl table = new CT_Tbl();
            table.AddNewTblPr().AddNewTblW().w = "1000";

            XWPFTable xtab = new XWPFTable(table, doc);

            Assert.AreEqual(1000, xtab.GetWidth());

            xtab.SetWidth(100);
            Assert.AreEqual(100, int.Parse(table.tblPr.tblW.w));
        }
        [TestMethod]
        public void TestSetHeight()
        {
            XWPFDocument doc = new XWPFDocument();

            CT_Tbl table = new CT_Tbl();

            XWPFTable xtab = new XWPFTable(table, doc);
            XWPFTableRow row = xtab.CreateRow();
            row.SetHeight(20);
            Assert.AreEqual(20, row.GetHeight());
        }

        [TestMethod]
        public void TestCreateTable()
        {
            // open an empty document
            XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");

            // create a table with 5 rows and 7 coloumns
            int noRows = 5;
            int noCols = 7;
            XWPFTable table = doc.CreateTable(noRows, noCols);

            // assert the table is empty
            List<XWPFTableRow> rows = table.GetRows();
            Assert.AreEqual(noRows, rows.Count, "Table has less rows than requested.");
            foreach (XWPFTableRow xwpfRow in rows)
            {
                Assert.IsNotNull(xwpfRow);
                for (int i = 0; i < 7; i++)
                {
                    XWPFTableCell xwpfCell = xwpfRow.GetCell(i);
                    Assert.IsNotNull(xwpfCell);
                    Assert.AreEqual(1, xwpfCell.Paragraphs.Count, "Empty cells should not have one paragraph.");
                    xwpfCell = xwpfRow.GetCell(i);
                    Assert.AreEqual(1, xwpfCell.Paragraphs.Count, "Calling 'getCell' must not modify cells content.");
                }
            }
            doc.Package.Revert();
        }
    }
}