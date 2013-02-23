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
    using NPOI.HSSF.Record;
    using NPOI.POIFS.FileSystem;
    using NPOI.HSSF.EventUserModel.DummyRecord;

    using NUnit.Framework;

    [TestFixture]
    public class TestMissingRecordAwareHSSFListener
    {

        private Record[] r;

        private void ReadRecords(String sampleFileName)
        {
            HSSFRequest req = new HSSFRequest();
            MockHSSFListener mockListen = new MockHSSFListener();
            MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(mockListen);
            req.AddListenerForAllRecords(listener);

            HSSFEventFactory factory = new HSSFEventFactory();
            try
            {
                Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);
                POIFSFileSystem fs = new POIFSFileSystem(is1);
                factory.ProcessWorkbookEvents(req, fs);
            }
            catch (IOException)
            {
                throw;
            }

            r = mockListen.GetRecords();
            Assert.IsTrue(r.Length > 100);
        }
        public void OpenNormal()
        {
            ReadRecords("MissingBits.xls");
        }
        [Test]
        public void TestMissingRowRecords()
        {
            OpenNormal();

            // We have rows 0, 1, 2, 20 and 21
            int row0 = -1;
            for (int i = 0; i < r.Length; i++)
            {
                if (r[i] is RowRecord)
                {
                    RowRecord rr = (RowRecord)r[i];
                    if (rr.RowNumber == 0) { row0 = i; }
                }
            }
            Assert.IsTrue(row0 > -1);

            // Following row 0, we should have 1, 2, then dummy, then 20+21+22
            Assert.IsTrue(r[row0] is RowRecord);
            Assert.IsTrue(r[row0 + 1] is RowRecord);
            Assert.IsTrue(r[row0 + 2] is RowRecord);
            Assert.IsTrue(r[row0 + 3] is MissingRowDummyRecord);
            Assert.IsTrue(r[row0 + 4] is MissingRowDummyRecord);
            Assert.IsTrue(r[row0 + 5] is MissingRowDummyRecord);
            Assert.IsTrue(r[row0 + 6] is MissingRowDummyRecord);
            // ...
            Assert.IsTrue(r[row0 + 18] is MissingRowDummyRecord);
            Assert.IsTrue(r[row0 + 19] is MissingRowDummyRecord);
            Assert.IsTrue(r[row0 + 20] is RowRecord);
            Assert.IsTrue(r[row0 + 21] is RowRecord);
            Assert.IsTrue(r[row0 + 22] is RowRecord);

            // Check things had the right row numbers
            RowRecord rr2;
            rr2 = (RowRecord)r[row0 + 2];
            Assert.AreEqual(2, rr2.RowNumber);
            rr2 = (RowRecord)r[row0 + 20];
            Assert.AreEqual(20, rr2.RowNumber);
            rr2 = (RowRecord)r[row0 + 21];
            Assert.AreEqual(21, rr2.RowNumber);

            MissingRowDummyRecord mr;
            mr = (MissingRowDummyRecord)r[row0 + 3];
            Assert.AreEqual(3, mr.RowNumber);
            mr = (MissingRowDummyRecord)r[row0 + 4];
            Assert.AreEqual(4, mr.RowNumber);
            mr = (MissingRowDummyRecord)r[row0 + 5];
            Assert.AreEqual(5, mr.RowNumber);
            mr = (MissingRowDummyRecord)r[row0 + 18];
            Assert.AreEqual(18, mr.RowNumber);
            mr = (MissingRowDummyRecord)r[row0 + 19];
            Assert.AreEqual(19, mr.RowNumber);
        }
        [Test]
        public void TestEndOfRowRecords()
        {
            OpenNormal();

            // Find the cell at 0,0
            int cell00 = -1;
            for (int i = 0; i < r.Length; i++)
            {
                if (r[i] is LabelSSTRecord)
                {
                    LabelSSTRecord lr = (LabelSSTRecord)r[i];
                    if (lr.Row == 0 && lr.Column == 0) { cell00 = i; }
                }
            }
            Assert.IsTrue(cell00 > -1);

            // We have rows 0, 1, 2, 20 and 21
            // Row 0 has 1 entry
            // Row 1 has 4 entries
            // Row 2 has 6 entries
            // Row 20 has 5 entries
            // Row 21 has 7 entries
            // Row 22 has 12 entries

            // Row 0
            Assert.IsFalse(r[cell00 + 0] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 1] is LastCellOfRowDummyRecord);
            // Row 1
            Assert.IsFalse(r[cell00 + 2] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 3] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 4] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 5] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 6] is LastCellOfRowDummyRecord);
            // Row 2
            Assert.IsFalse(r[cell00 + 7] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 8] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 9] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 10] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 11] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 12] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 13] is LastCellOfRowDummyRecord);
            // Row 3 -> 19
            Assert.IsTrue(r[cell00 + 14] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 15] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 16] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 17] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 18] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 19] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 20] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 21] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 22] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 23] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 24] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 25] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 26] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 27] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 28] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 29] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 30] is LastCellOfRowDummyRecord);
            // Row 20
            Assert.IsFalse(r[cell00 + 31] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 32] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 33] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 34] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 35] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 36] is LastCellOfRowDummyRecord);
            // Row 21
            Assert.IsFalse(r[cell00 + 37] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 38] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 39] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 40] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 41] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 42] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 43] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 44] is LastCellOfRowDummyRecord);
            // Row 22
            Assert.IsFalse(r[cell00 + 45] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 46] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 47] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 48] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 49] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 50] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 51] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 52] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 53] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 54] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 55] is LastCellOfRowDummyRecord);
            Assert.IsFalse(r[cell00 + 56] is LastCellOfRowDummyRecord);
            Assert.IsTrue(r[cell00 + 57] is LastCellOfRowDummyRecord);

            // Check the numbers of the last seen columns
            LastCellOfRowDummyRecord[] lrs = new LastCellOfRowDummyRecord[24];
            int lrscount = 0;
            for (int i = 0; i < r.Length; i++)
            {
                if (r[i] is LastCellOfRowDummyRecord)
                {
                    lrs[lrscount] = (LastCellOfRowDummyRecord)r[i];
                    lrscount++;
                }
            }

            Assert.AreEqual(0, lrs[0].LastColumnNumber);
            Assert.AreEqual(0, lrs[0].Row);

            Assert.AreEqual(3, lrs[1].LastColumnNumber);
            Assert.AreEqual(1, lrs[1].Row);

            Assert.AreEqual(5, lrs[2].LastColumnNumber);
            Assert.AreEqual(2, lrs[2].Row);

            for (int i = 3; i <= 19; i++)
            {
                Assert.AreEqual(-1, lrs[i].LastColumnNumber);
                Assert.AreEqual(i, lrs[i].Row);
            }

            Assert.AreEqual(4, lrs[20].LastColumnNumber);
            Assert.AreEqual(20, lrs[20].Row);

            Assert.AreEqual(6, lrs[21].LastColumnNumber);
            Assert.AreEqual(21, lrs[21].Row);

            Assert.AreEqual(11, lrs[22].LastColumnNumber);
            Assert.AreEqual(22, lrs[22].Row);
        }

        [Test]
        public void TestMissingCellRecords()
        {
            OpenNormal();

            // Find the cell at 0,0
            int cell00 = -1;
            for (int i = 0; i < r.Length; i++)
            {
                if (r[i] is LabelSSTRecord)
                {
                    LabelSSTRecord lr = (LabelSSTRecord)r[i];
                    if (lr.Row == 0 && lr.Column == 0) { cell00 = i; }
                }
            }
            Assert.IsTrue(cell00 > -1);

            // We have rows 0, 1, 2, 20 and 21
            // Row 0 has 1 entry, 0
            // Row 1 has 4 entries, 0+3
            // Row 2 has 6 entries, 0+5
            // Row 20 has 5 entries, 0-5
            // Row 21 has 7 entries, 0+1+3+5+6
            // Row 22 has 12 entries, 0+3+4+11

            // Row 0
            Assert.IsFalse(r[cell00 + 0] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 1] is MissingCellDummyRecord);

            // Row 1
            Assert.IsFalse(r[cell00 + 2] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 3] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 4] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 5] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 6] is MissingCellDummyRecord);

            // Row 2
            Assert.IsFalse(r[cell00 + 7] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 8] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 9] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 10] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 11] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 12] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 13] is MissingCellDummyRecord);

            // Row 3-19
            Assert.IsFalse(r[cell00 + 14] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 15] is MissingCellDummyRecord);

            // Row 20
            Assert.IsFalse(r[cell00 + 31] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 32] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 33] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 34] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 35] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 36] is MissingCellDummyRecord);

            // Row 21
            Assert.IsFalse(r[cell00 + 37] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 38] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 39] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 40] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 41] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 42] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 43] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 44] is MissingCellDummyRecord);

            // Row 22
            Assert.IsFalse(r[cell00 + 45] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 46] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 47] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 48] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 49] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 50] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 51] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 52] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 53] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 54] is MissingCellDummyRecord);
            Assert.IsTrue(r[cell00 + 55] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 56] is MissingCellDummyRecord);
            Assert.IsFalse(r[cell00 + 57] is MissingCellDummyRecord);

            // Check some numbers
            MissingCellDummyRecord mc;

            mc = (MissingCellDummyRecord)r[cell00 + 3];
            Assert.AreEqual(1, mc.Row);
            Assert.AreEqual(1, mc.Column);
            mc = (MissingCellDummyRecord)r[cell00 + 4];
            Assert.AreEqual(1, mc.Row);
            Assert.AreEqual(2, mc.Column);

            mc = (MissingCellDummyRecord)r[cell00 + 8];
            Assert.AreEqual(2, mc.Row);
            Assert.AreEqual(1, mc.Column);
            mc = (MissingCellDummyRecord)r[cell00 + 9];
            Assert.AreEqual(2, mc.Row);
            Assert.AreEqual(2, mc.Column);

            mc = (MissingCellDummyRecord)r[cell00 + 55];
            Assert.AreEqual(22, mc.Row);
            Assert.AreEqual(10, mc.Column);
        }

        // Make sure we don't put in any extra new lines
        //  that aren't alReady there
        [Test]
        public void TestNoExtraNewLines()
        {
            // Load a different file
            // This file has has something in lines 1-33
            ReadRecords("MRExtraLines.xls");

            int rowCount = 0;
            for (int i = 0; i < r.Length; i++)
            {
                if (r[i] is LastCellOfRowDummyRecord)
                {
                    LastCellOfRowDummyRecord eor = (LastCellOfRowDummyRecord)r[i];
                    Assert.AreEqual(rowCount, eor.Row);
                    rowCount++;
                }
            }
            // Check we got the 33 rows
            Assert.AreEqual(33, rowCount);
        }

        private class MockHSSFListener : IHSSFListener
        {
            public MockHSSFListener() { }
            private ArrayList _records = new ArrayList();
            private bool logToStdOut = false;

            public void ProcessRecord(Record record)
            {
                _records.Add(record);

                if (record is MissingRowDummyRecord)
                {
                    MissingRowDummyRecord mr = (MissingRowDummyRecord)record;
                    log("Got dummy row " + mr.RowNumber);
                }
                if (record is MissingCellDummyRecord)
                {
                    MissingCellDummyRecord mc = (MissingCellDummyRecord)record;
                    log("Got dummy cell " + mc.Row + " " + mc.Column);
                }
                if (record is LastCellOfRowDummyRecord)
                {
                    LastCellOfRowDummyRecord lc = (LastCellOfRowDummyRecord)record;
                    log("Got end-of row, row was " + lc.Row + ", last column was " + lc.LastColumnNumber);
                }

                if (record is BOFRecord)
                {
                    BOFRecord r = (BOFRecord)record;
                    if (r.Type == BOFRecordType.Worksheet)
                    {
                        log("On new sheet");
                    }
                }
                if (record is RowRecord)
                {
                    RowRecord rr = (RowRecord)record;
                    log("Starting row #" + rr.RowNumber);
                }
            }
            private void log(String msg)
            {
                if (logToStdOut)
                {
                    Console.WriteLine(msg);
                }
            }
            public Record[] GetRecords()
            {
                Record[] result = (Record[])_records.ToArray(typeof(Record));
                return result;
            }
        }

        /**
         * Make sure that the presence of shared formulas does not cause extra 
         * end-of-row records.
         */
        [Test]
        public void TestEndOfRow_bug45672()
        {
            ReadRecords("ex45672.xls");
            Record[] rr = r;
            int eorCount = 0;
            int sfrCount = 0;
            for (int i = 0; i < rr.Length; i++)
            {
                Record record = rr[i];
                if (record is SharedFormulaRecord)
                {
                    sfrCount++;
                }
                if (record is LastCellOfRowDummyRecord)
                {
                    eorCount++;
                }
            }
            if (eorCount == 2)
            {
                throw new AssertionException("Identified bug 45672");
            }
            Assert.AreEqual(1, eorCount);
            Assert.AreEqual(1, sfrCount);
        }

        /**
	     * MulBlank records hold multiple blank cells. Check we
	     *  can handle them correctly.
	     */
        [Test]
        public void TestMulBlankHandling()
        {
            ReadRecords("45672.xls");

            // Check that we don't have any MulBlankRecords, but do
            //  have lots of BlankRecords
            Record[] rr = r;
            int eorCount = 0;
            int mbrCount = 0;
            int brCount = 0;
            for (int i = 0; i < rr.Length; i++)
            {
                Record record = rr[i];
                if (record is MulBlankRecord)
                {
                    mbrCount++;
                }
                if (record is BlankRecord)
                {
                    brCount++;
                }
                if (record is LastCellOfRowDummyRecord)
                {
                    eorCount++;
                }
            }
            if (mbrCount > 0)
            {
                throw new AssertionException("Identified bug 45672");
            }
            if (brCount < 20)
            {
                throw new AssertionException("Identified bug 45672");
            }
            if (eorCount != 2)
            {
                throw new AssertionException("Identified bug 45672");
            }
            Assert.AreEqual(2, eorCount);
        }
        [Test]
        public void TestStringRecordHandling()
        {
            ReadRecords("53588.xls");
            Record[] rr = r;
            int missingCount = 0;
            int lastCount = 0;
            for (int i = 0; i < rr.Length; i++)
            {
                Record record = rr[i];
                if (record is MissingCellDummyRecord)
                {
                    missingCount++;
                }
                if (record is LastCellOfRowDummyRecord)
                {
                    lastCount++;
                }
            }
            Assert.AreEqual(1, missingCount);
            Assert.AreEqual(1, lastCount);
        }
    }
}