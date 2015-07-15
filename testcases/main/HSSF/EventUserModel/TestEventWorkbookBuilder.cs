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

namespace TestCases.HSSF.EventUserModel
{
    using System;
    using System.IO;
    using System.Collections;

    using NPOI.HSSF;
    using NPOI.HSSF.EventUserModel;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Record;
    using NPOI.POIFS.FileSystem;
    using NPOI.SS.Formula;

    using NUnit.Framework;
    using NPOI.SS.Formula.PTG;


    [TestFixture]
    public class TestEventWorkbookBuilder
    {
        private MockHSSFListener mockListen;
        private EventWorkbookBuilder.SheetRecordCollectingListener listener;

        [SetUp]
        public void SetUp()
        {
            HSSFRequest req = new HSSFRequest();
            mockListen = new MockHSSFListener();
            listener = new EventWorkbookBuilder.SheetRecordCollectingListener(mockListen);
            req.AddListenerForAllRecords(listener);

            HSSFEventFactory factory = new HSSFEventFactory();
            try
            {
                Stream is1 = HSSFTestDataSamples.OpenSampleFileStream("3dFormulas.xls");
                POIFSFileSystem fs = new POIFSFileSystem(is1);
                factory.ProcessWorkbookEvents(req, fs);
            }
            catch (IOException)
            {
                throw;
            }
        }
        [Test]
        public void TestBasics()
        {
            Assert.IsNotNull(listener.GetSSTRecord());
            Assert.IsNotNull(listener.GetBoundSheetRecords());
            Assert.IsNotNull(listener.GetExternSheetRecords());
        }
        [Test]
        public void TestGetStubWorkbooks()
        {
            Assert.IsNotNull(listener.GetStubWorkbook());
            Assert.IsNotNull(listener.GetStubHSSFWorkbook());
        }
        [Test]
        public void TestContents()
        {
            Assert.AreEqual(2, listener.GetSSTRecord().NumStrings);
            Assert.AreEqual(3, listener.GetBoundSheetRecords().Length);
            Assert.AreEqual(1, listener.GetExternSheetRecords().Length);

            Assert.AreEqual(3, listener.GetStubWorkbook().NumSheets);

            InternalWorkbook ref1 = listener.GetStubWorkbook();
            Assert.AreEqual("Sh3", ref1.FindSheetFirstNameFromExternSheet(0));
            Assert.AreEqual("Sheet1", ref1.FindSheetFirstNameFromExternSheet(1));
            Assert.AreEqual("S2", ref1.FindSheetFirstNameFromExternSheet(2));
        }
        [Test]
        public void TestFormulas()
        {

            FormulaRecord[] fRecs = mockListen.GetFormulaRecords();

            // Check our formula records
            Assert.AreEqual(6, fRecs.Length);

            InternalWorkbook stubWB = listener.GetStubWorkbook();
            Assert.IsNotNull(stubWB);
            HSSFWorkbook stubHSSF = listener.GetStubHSSFWorkbook();
            Assert.IsNotNull(stubHSSF);

            // Check these stubs have the right stuff on them
            Assert.AreEqual("Sheet1", stubWB.GetSheetName(0));
            Assert.AreEqual("S2", stubWB.GetSheetName(1));
            Assert.AreEqual("Sh3", stubWB.GetSheetName(2));

            // Check we can Get the formula without breaking
            for (int i = 0; i < fRecs.Length; i++)
            {
                HSSFFormulaParser.ToFormulaString(stubHSSF, fRecs[i].ParsedExpression);
            }

            // Peer into just one formula, and check that
            //  all the ptgs give back the right things
            Ptg[] ptgs = fRecs[0].ParsedExpression;
            Assert.AreEqual(1, ptgs.Length);
            Assert.IsTrue(ptgs[0] is Ref3DPtg);

            Ref3DPtg ptg = (Ref3DPtg)ptgs[0];
            HSSFEvaluationWorkbook book = HSSFEvaluationWorkbook.Create(stubHSSF);
            Assert.AreEqual("Sheet1!A1", ptg.ToFormulaString(book));


            // Now check we Get the right formula back for
            //  a few sample ones
            FormulaRecord fr;

            // Sheet 1 A2 is on same sheet
            fr = fRecs[0];
            Assert.AreEqual(1, fr.Row);
            Assert.AreEqual(0, fr.Column);
            Assert.AreEqual("Sheet1!A1", HSSFFormulaParser.ToFormulaString(stubHSSF, fr.ParsedExpression));

            // Sheet 1 A5 is to another sheet
            fr = fRecs[3];
            Assert.AreEqual(4, fr.Row);
            Assert.AreEqual(0, fr.Column);
            Assert.AreEqual("'S2'!A1", HSSFFormulaParser.ToFormulaString(stubHSSF, fr.ParsedExpression));

            // Sheet 1 A7 is to another sheet, range
            fr = fRecs[5];
            Assert.AreEqual(6, fr.Row);
            Assert.AreEqual(0, fr.Column);
            Assert.AreEqual("SUM(Sh3!A1:A4)", HSSFFormulaParser.ToFormulaString(stubHSSF, fr.ParsedExpression));


            // Now, load via Usermodel and re-check
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("3dFormulas.xls");
            Assert.AreEqual("Sheet1!A1", wb.GetSheetAt(0).GetRow(1).GetCell(0).CellFormula);
            Assert.AreEqual("SUM(Sh3!A1:A4)", wb.GetSheetAt(0).GetRow(6).GetCell(0).CellFormula);
        }

        private class MockHSSFListener : IHSSFListener
        {
            public MockHSSFListener() { }
            private ArrayList _records = new ArrayList();
            private ArrayList _frecs = new ArrayList();

            public void ProcessRecord(Record record)
            {
                _records.Add(record);
                if (record is FormulaRecord)
                {
                    _frecs.Add(record);
                }
            }
            public FormulaRecord[] GetFormulaRecords()
            {
                FormulaRecord[] result = (FormulaRecord[])_frecs.ToArray(typeof(FormulaRecord));
                return result;
            }
        }
    }
}