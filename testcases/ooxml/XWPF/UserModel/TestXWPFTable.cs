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
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;
using org.Openxmlformats.schemas.wordProcessingml.x2006.main;

/**
 * Tests for XWPF Run
 */
    [TestClass]
    public class TestXWPFTable 
{

    protected void SetUp() {
        /*
          XWPFDocument doc = new XWPFDocument();
          p = doc.CreateParagraph();

          this.ctRun = CTR.Factory.NewInstance();
       */
    }

        [TestMethod]
    public void TestConstructor(){
        XWPFDocument doc = new XWPFDocument();
        CTTbl ctTable = CTTbl.Factory.NewInstance();
        XWPFTable xtab = new XWPFTable(ctTable, doc);
        Assert.IsNotNull(xtab);
        Assert.AreEqual(1, ctTable.SizeOfTrArray());
        Assert.AreEqual(1, ctTable.GetTrArray(0).sizeOfTcArray());
        Assert.IsNotNull(ctTable.GetTrArray(0).GetTcArray(0).GetPArray(0));

        ctTable = CTTbl.Factory.NewInstance();
        xtab = new XWPFTable(ctTable, doc, 3, 2);
        Assert.IsNotNull(xtab);
        Assert.AreEqual(3, ctTable.SizeOfTrArray());
        Assert.AreEqual(2, ctTable.GetTrArray(0).sizeOfTcArray());
        Assert.IsNotNull(ctTable.GetTrArray(0).GetTcArray(0).GetPArray(0));
    }


        [TestMethod]
    public void TestGetText(){
        XWPFDocument doc = new XWPFDocument();
        CTTbl table = CTTbl.Factory.NewInstance();
        CTRow row = table.AddNewTr();
        CTTc cell = row.AddNewTc();
        CTP paragraph = cell.AddNewP();
        CTR run = paragraph.AddNewR();
        CTText text = Run.AddNewT();
        text.StringValue=("finally I can Write!");

        XWPFTable xtab = new XWPFTable(table, doc);
        Assert.AreEqual("finally I can Write!\n", xtab.Text);
    }


        [TestMethod]
    public void TestCreateRow(){
        XWPFDocument doc = new XWPFDocument();

        CTTbl table = CTTbl.Factory.NewInstance();
        CTRow r1 = table.AddNewTr();
        r1.AddNewTc().AddNewP();
        r1.AddNewTc().AddNewP();
        CTRow r2 = table.AddNewTr();
        r2.AddNewTc().AddNewP();
        r2.AddNewTc().AddNewP();
        CTRow r3 = table.AddNewTr();
        r3.AddNewTc().AddNewP();
        r3.AddNewTc().AddNewP();

        XWPFTable xtab = new XWPFTable(table, doc);
        Assert.AreEqual(3, xtab.NumberOfRows);
        Assert.IsNotNull(xtab.GetRow(2));

        //add a new row
        xtab.CreateRow();
        Assert.AreEqual(4, xtab.NumberOfRows);

        //check number of cols
        Assert.AreEqual(2, table.GetTrArray(0).sizeOfTcArray());

        //check creation of first row
        xtab = new XWPFTable(CTTbl.Factory.NewInstance(), doc);
        Assert.AreEqual(1, xtab.CTTbl.GetTrArray(0).sizeOfTcArray());
    }


    public void testSetWidth {
        XWPFDocument doc = new XWPFDocument();
        
        CTTbl table = CTTbl.Factory.NewInstance();
        table.AddNewTblPr().AddNewTblW().W=(new Bigint("1000"));

        XWPFTable xtab = new XWPFTable(table, doc);

        Assert.AreEqual(1000, xtab.Width);

        xtab.Width=(100);
        Assert.AreEqual(100, table.TblPr.TblW.W.IntValue());
    }

    public void testSetHeight {
        XWPFDocument doc = new XWPFDocument();

        CTTbl table = CTTbl.Factory.NewInstance();

        XWPFTable xtab = new XWPFTable(table, doc);
        XWPFTableRow row = xtab.CreateRow();
        row.Height=(20);
        Assert.AreEqual(20, row.Height);
    }

        [TestMethod]
    public void TestCreateTable(){
       // open an empty document
       XWPFDocument doc = XWPFTestDataSamples.OpenSampleDocument("sample.docx");

       // create a table with 5 rows and 7 coloumns
       int noRows = 5; 
       int noCols = 7;
       XWPFTable table = doc.CreateTable(noRows,noCols);

       // assert the table is empty
       List<XWPFTableRow> rows = table.Rows;
       Assert.AreEqual("Table has less rows than requested.", noRows, rows.Size());
       foreach (XWPFTableRow xwpfRow in rows)
       {
          Assert.IsNotNull(xwpfRow);
          for (int i = 0 ; i < 7 ; i++)
          {
             XWPFTableCell xwpfCell = xwpfRow.GetCell(i);
             Assert.IsNotNull(xwpfCell);
             Assert.AreEqual("Empty cells should not have one paragraph.",1,xwpfCell.Paragraphs.Size());
             xwpfCell = xwpfRow.GetCell(i);
             Assert.AreEqual("Calling 'getCell' must not modify cells content.",1,xwpfCell.Paragraphs.Size());
          }
       }
       doc.Package.Revert();
    }
}
