/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.Util;
using NUnit.Framework;
using TestCases.HSSF.UserModel;
using NPOI.HSSF.Model;
using NPOI.DDF;
using NPOI.HSSF.Record.Aggregates;

namespace TestCases.HSSF.Model
{
    /**
 * @author Yegor Kozlov
 * @author Evgeniy Berlog
 */
    [TestFixture]
    public class TestDrawingAggregate
    {
        /**
         *  information about Drawing aggregate in a worksheet
         */
        private class DrawingAggregateInfo
        {
            /**
             * start and end indices of the aggregate in the worksheet stream
             */
            private int startRecordIndex, endRecordIndex;
            /**
             * the records being aggregated
             */
            private List<RecordBase> aggRecords;

            /**
             * @return aggregate info or null if the sheet does not contain Drawing objects
             */
            internal static DrawingAggregateInfo Get(HSSFSheet sheet)
            {
                DrawingAggregateInfo info = null;
                InternalSheet isheet = HSSFTestHelper.GetSheetForTest(sheet);
                List<RecordBase> records = isheet.Records;
                for (int i = 0; i < records.Count; i++)
                {
                    RecordBase rb = records[(i)];
                    if ((rb is DrawingRecord) && info == null)
                    {
                        info = new DrawingAggregateInfo();
                        info.startRecordIndex = i;
                        info.endRecordIndex = i;
                    }
                    else if (info != null && (
                          rb is DrawingRecord
                                  || rb is ObjRecord
                                  || rb is TextObjectRecord
                                  || rb is ContinueRecord
                                  || rb is NoteRecord
                  ))
                    {
                        info.endRecordIndex = i;
                    }
                    else
                    {
                        if (rb is EscherAggregate)
                            throw new InvalidOperationException("Drawing data already aggregated. " +
                                    "You should cal this method before the first invocation of HSSFSheet#getDrawingPatriarch()");
                        if (info != null) break;
                    }
                }
                if (info != null)
                {
                    info.aggRecords = new List<RecordBase>(
                            records.GetRange(info.startRecordIndex, info.endRecordIndex + 1));
                }
                return info;
            }

            /**
             * @return the raw data being aggregated
             */
            internal byte[] GetRawBytes()
            {
                MemoryStream out1 = new MemoryStream();
                foreach (RecordBase rb in aggRecords)
                {
                    NPOI.HSSF.Record.Record r = (NPOI.HSSF.Record.Record)rb;
                    try
                    {
                        byte[] data = r.Serialize();
                        out1.Write(data, 0, data.Length);
                    }
                    catch (IOException e)
                    {
                        throw new RuntimeException(e);
                    }
                }
                return out1.ToArray();
            }
        }

        /**
         * iterate over all sheets, aggregate Drawing records (if there are any)
         * and remember information about the aggregated data.
         * Then serialize the workbook, read back and assert that the aggregated data is preserved.
         *
         * The assertion is strict meaning that the Drawing data before and After save must be Equal.
         */
        private static void assertWriteAndReadBack(HSSFWorkbook wb)
        {
            // map aggregate info by sheet index
            Dictionary<int, DrawingAggregateInfo> aggs = new Dictionary<int, DrawingAggregateInfo>();
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                HSSFSheet sheet = wb.GetSheetAt(i) as HSSFSheet;
                DrawingAggregateInfo info = DrawingAggregateInfo.Get(sheet);
                if (info != null)
                {
                    aggs.Add(i, info);
                    HSSFPatriarch p = sheet.DrawingPatriarch as HSSFPatriarch;

                    // compare aggregate.Serialize() with raw bytes from the record stream
                    EscherAggregate agg = HSSFTestHelper.GetEscherAggregate(p);

                    byte[] dgBytes1 = info.GetRawBytes();
                    byte[] dgBytes2 = agg.Serialize();

                    Assert.AreEqual(dgBytes1.Length, dgBytes2.Length, "different size of raw data ande aggregate.Serialize()");
                    Assert.IsTrue(Arrays.Equals(dgBytes1, dgBytes2), "raw Drawing data (" + dgBytes1.Length + " bytes) and aggregate.Serialize() are different.");
                }
            }

            if (aggs.Count != 0)
            {
                HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb);
                for (int i = 0; i < wb2.NumberOfSheets; i++)
                {
                    DrawingAggregateInfo info1 = aggs[(i)];
                    if (info1 != null)
                    {
                        HSSFSheet sheet2 = wb2.GetSheetAt(i) as HSSFSheet;
                        DrawingAggregateInfo info2 = DrawingAggregateInfo.Get(sheet2);
                        byte[] dgBytes1 = info1.GetRawBytes();
                        byte[] dgBytes2 = info2.GetRawBytes();
                        Assert.AreEqual(dgBytes1.Length, dgBytes2.Length, "different size of Drawing data before and After save");
                        Assert.IsTrue(Arrays.Equals(dgBytes1, dgBytes2), "drawing data (" + dgBytes1.Length + " bytes) before and After save is different.");
                    }
                }
            }
        }

        /**
         * Test that we correctly read and write Drawing aggregates
         *  in all .xls files in POI Test samples
         */
        [Ignore("refactor it")]
        [Test]
        public void TestAllTestSamples()
        {
            //File[] xls = new File(System.GetProperty("POI.testdata.path"), "spreadsheet").listFiles(
            //        new FilenameFilter() {
            //            public bool accept(File dir, String name) {
            //                return name.EndsWith(".xls");
            //            }
            //        }
            //);
            //Assert.IsNotNull(xls,
            //    "Need to find files in test-data path, had path: " + new File(System.getProperty("POI.testdata.path"), "spreadsheet"));
            //foreach(File file in xls) {
            //    HSSFWorkbook wb;
            //    try {
            //        wb = HSSFTestDataSamples.OpenSampleWorkbook(file.GetName());
            //    } catch (Throwable e){
            //        // don't bother about files we cannot read - they are different bugs
            //        // Console.WriteLine("[WARN]  Cannot read " + file.GetName());
            //        continue;
            //    }
            //    assertWriteAndReadBack(wb);
            //}
        }

        /**
         * when Reading incomplete data ensure that the Serialized bytes
         match the source
         */
        [Test]
        public void TestIncompleteData()
        {
            //EscherDgContainer and EscherSpgrContainer length exceeds the actual length of the data
            String hex =
                    " 0F 00 02 F0 30 03 00 00 10 00 08 F0 08 00 00 " +
                    " 00 07 00 00 00 B2 04 00 00 0F 00 03 F0 18 03 00 " +
                    " 00 0F 00 04 F0 28 00 00 00 01 00 09 F0 10 00 00 " +
                    " 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                    " 00 02 00 0A F0 08 00 00 00 00 04 00 00 05 00 00 " +
                    " 00 0F 00 04 F0 74 00 00 00 92 0C 0A F0 08 00 00 " +
                    " 00 AD 04 00 00 00 0A 00 00 63 00 0B F0 3A 00 00 " +
                    " 00 7F 00 04 01 E5 01 BF 00 08 00 08 00 81 01 4E " +
                    " 00 00 08 BF 01 10 00 10 00 80 C3 16 00 00 00 BF " +
                    " 03 00 00 02 00 44 00 69 00 61 00 67 00 72 00 61 " +
                    " 00 6D 00 6D 00 20 00 32 00 00 00 00 00 10 F0 12 " +
                    " 00 00 00 00 00 05 00 00 00 01 00 00 00 0B 00 00 " +
                    " 00 0F 00 66 00 00 00 11 F0 00 00 00 00 ";
            byte[] buffer = HexRead.ReadFromString(hex);

            List<EscherRecord> records = new List<EscherRecord>();
            IEscherRecordFactory recordFactory = new DefaultEscherRecordFactory();
            int pos = 0;
            while (pos < buffer.Length)
            {
                EscherRecord r = recordFactory.CreateRecord(buffer, pos);
                int bytesRead = r.FillFields(buffer, pos, recordFactory);
                records.Add(r);
                pos += bytesRead;
            }
            Assert.AreEqual(buffer.Length, pos, "data was not fully Read");

            // serialize to byte array
            MemoryStream out1 = new MemoryStream();
            try
            {
                foreach (EscherRecord r in records)
                {
                    byte[] data = r.Serialize();
                    out1.Write(data, 0, data.Length);
                }
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
            Assert.AreEqual(HexDump.ToHex(buffer, 10), HexDump.ToHex(out1.ToArray(), 10));
        }

        /**
         * TODO: figure out why it fails with "RecordFormatException: 0 bytes written but GetRecordSize() reports 80"
         */
        [Test]
        public void TestFailing()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("15573.xls");
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;
            HSSFPatriarch p = sh.DrawingPatriarch as HSSFPatriarch;

            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
        }

        private static byte[] ToArray(List<RecordBase> records)
        {
            MemoryStream out1 = new MemoryStream();
            foreach (RecordBase rb in records)
            {
                NPOI.HSSF.Record.Record r = (NPOI.HSSF.Record.Record)rb;
                try
                {
                    byte[] data = r.Serialize();
                    out1.Write(data, 0, data.Length);
                }
                catch (IOException e)
                {
                    throw new RuntimeException(e);
                }
            }
            return out1.ToArray();
        }
        [Test]
        public void TestSolverContainerMustBeSavedDuringSerialization()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("SolverContainerAfterSPGR.xls");
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;
            InternalSheet ish = HSSFTestHelper.GetSheetForTest(sh);
            List<RecordBase> records = ish.Records;
            // records to be aggregated
            List<RecordBase> dgRecords = records.GetRange(19, 22-19);
            byte[] dgBytes = ToArray(dgRecords);
            HSSFPatriarch p = sh.DrawingPatriarch as HSSFPatriarch;
            EscherAggregate agg = (EscherAggregate)ish.FindFirstRecordBySid(EscherAggregate.sid);
            Assert.AreEqual(agg.EscherRecords[0].ChildRecords.Count, 3);
            Assert.AreEqual(agg.EscherRecords[0].GetChild(2).RecordId, EscherContainerRecord.SOLVER_CONTAINER);
            wb = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            sh = wb.GetSheetAt(0) as HSSFSheet;
            p = sh.DrawingPatriarch as HSSFPatriarch;
            ish = HSSFTestHelper.GetSheetForTest(sh);
            agg = (EscherAggregate)ish.FindFirstRecordBySid(EscherAggregate.sid);
            Assert.AreEqual(agg.EscherRecords[0].ChildRecords.Count, 3);
            Assert.AreEqual(agg.EscherRecords[0].GetChild(2).RecordId, EscherContainerRecord.SOLVER_CONTAINER);


            // collect Drawing records into a byte buffer.
            agg = (EscherAggregate)ish.FindFirstRecordBySid(EscherAggregate.sid);
            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length, "different size of Drawing data before and After save");
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data before and After save is different");
        }
        [Test]
        public void TestFileWithTextbox()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("text.xls");
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;
            InternalSheet ish = HSSFTestHelper.GetSheetForTest(sh);
            List<RecordBase> records = ish.Records;
            // records to be aggregated
            List<RecordBase> dgRecords = records.GetRange(19, 23-19);
            byte[] dgBytes = ToArray(dgRecords);
            HSSFPatriarch p = sh.DrawingPatriarch as HSSFPatriarch;

            // collect Drawing records into a byte buffer.
            EscherAggregate agg = (EscherAggregate)ish.FindFirstRecordBySid(EscherAggregate.sid);
            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length, "different size of Drawing data before and After save");
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data before and After save is different");
        }
        [Test]
        public void TestFileWithCharts()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("49581.xls");
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;
            InternalSheet ish = HSSFTestHelper.GetSheetForTest(sh);
            List<RecordBase> records = ish.Records;
            // records to be aggregated
            List<RecordBase> dgRecords = records.GetRange(19, 21-19);
            byte[] dgBytes = ToArray(dgRecords);
            HSSFPatriarch p = sh.DrawingPatriarch as HSSFPatriarch;

            // collect Drawing records into a byte buffer.
            EscherAggregate agg = (EscherAggregate)ish.FindFirstRecordBySid(EscherAggregate.sid);
            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length, "different size of Drawing data before and After save");
            for (int i = 0; i < dgBytes.Length; i++)
            {
                if (dgBytes[i] != dgBytesAfterSave[i])
                {
                    Console.WriteLine("pos = " + i);
                }
            }
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data before and After save is different");
        }

        /**
         * Test Reading Drawing aggregate from a Test file from Bugzilla 45129
         */
        [Test]
        public void Test45129()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("45129.xls");
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;

            InternalWorkbook iworkbook = HSSFTestHelper.GetWorkbookForTest(wb);
            InternalSheet isheet = HSSFTestHelper.GetSheetForTest(sh);

            List<RecordBase> records = isheet.Records;

            // the sheet's Drawing is not aggregated
            Assert.AreEqual(394, records.Count, "wrong size of sheet records stream");
            // the last record before the Drawing block
            Assert.IsTrue(records[(18)] is RowRecordsAggregate,
                "records.Get(18) is expected to be RowRecordsAggregate but was " + records[(18)].GetType().Name);

            // records to be aggregated
            List<RecordBase> dgRecords = records.GetRange(19, 389-19);
            // collect Drawing records into a byte buffer.
            byte[] dgBytes = ToArray(dgRecords);

            foreach (RecordBase rb in dgRecords)
            {
                NPOI.HSSF.Record.Record r = (NPOI.HSSF.Record.Record)rb;
                short sid = r.Sid;
                // we expect that Drawing block consists of either
                // DrawingRecord or ContinueRecord or ObjRecord or TextObjectRecord
                Assert.IsTrue(
                        sid == DrawingRecord.sid ||
                                sid == ContinueRecord.sid ||
                                sid == ObjRecord.sid ||
                                sid == TextObjectRecord.sid);
            }

            // the first record After the Drawing block
            Assert.IsTrue(records[(389)] is WindowTwoRecord, "records.Get(389) is expected to be Window2");

            // aggregate Drawing records.
            // The subrange [19, 388] is expected to be Replaced with a EscherAggregate object
            DrawingManager2 drawingManager = iworkbook.FindDrawingGroup();
            int loc = isheet.AggregateDrawingRecords(drawingManager, false);
            EscherAggregate agg = (EscherAggregate)records[(loc)];

            Assert.AreEqual(25, records.Count, "wrong size of the aggregated sheet records stream");
            Assert.IsTrue(records[18] is RowRecordsAggregate,
                "records.Get(18) is expected to be RowRecordsAggregate but was " + records[18].GetType().Name);
            Assert.IsTrue(records[19] is EscherAggregate,
                "records.Get(19) is expected to be EscherAggregate but was " + records[19].GetType().Name);
            Assert.IsTrue(records[20] is WindowTwoRecord,
                "records.Get(20) is expected to be Window2 but was " + records[20].GetType().Name);

            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length, "different size of Drawing data before and After save");
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data before and After save is different");
        }

        /**
         * Try to check file with such record sequence
         * ...
         * DrawingRecord
         * ContinueRecord
         * ObjRecord | TextObjRecord
         * ...
         */
        public void TestSerializeDrawingBigger8k()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("DrawingContinue.xls");
            InternalWorkbook iworkbook = HSSFTestHelper.GetWorkbookForTest(wb);
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;
            InternalSheet isheet = HSSFTestHelper.GetSheetForTest(sh);


            List<RecordBase> records = isheet.Records;

            // the sheet's Drawing is not aggregated
            Assert.AreEqual(32, records.Count, "wrong size of sheet records stream");
            // the last record before the Drawing block
            Assert.IsTrue(records[18] is RowRecordsAggregate,
                "records.Get(18) is expected to be RowRecordsAggregate but was " + records[18].GetType().Name);

            // records to be aggregated
            List<RecordBase> dgRecords = records.GetRange(19, 26-19);
            foreach (RecordBase rb in dgRecords)
            {
                NPOI.HSSF.Record.Record r = (NPOI.HSSF.Record.Record)rb;
                short sid = r.Sid;
                // we expect that Drawing block consists of either
                // DrawingRecord or ContinueRecord or ObjRecord or TextObjectRecord
                Assert.IsTrue(
                        sid == DrawingRecord.sid ||
                                sid == ContinueRecord.sid ||
                                sid == ObjRecord.sid ||
                                sid == NoteRecord.sid ||
                                sid == TextObjectRecord.sid);
            }
            // collect Drawing records into a byte buffer.
            byte[] dgBytes = ToArray(dgRecords);

            // the first record After the Drawing block
            Assert.IsTrue(records[(26)] is WindowTwoRecord, "records.Get(26) is expected to be Window2");

            // aggregate Drawing records.
            // The subrange [19, 38] is expected to be Replaced with a EscherAggregate object
            DrawingManager2 drawingManager = iworkbook.FindDrawingGroup();
            int loc = isheet.AggregateDrawingRecords(drawingManager, false);
            EscherAggregate agg = (EscherAggregate)records[(loc)];

            Assert.AreEqual(26, records.Count, "wrong size of the aggregated sheet records stream");
            Assert.IsTrue(records[(18)] is RowRecordsAggregate, 
                "records.Get(18) is expected to be RowRecordsAggregate but was " + records[(18)].GetType().Name);
            Assert.IsTrue(records[(19)] is EscherAggregate,
                "records.Get(19) is expected to be EscherAggregate but was " + records[19].GetType().Name);
            Assert.IsTrue(records[(20)] is WindowTwoRecord,
                "records.Get(20) is expected to be Window2 but was " + records[20].GetType().Name);

            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length,"different size of Drawing data before and After save");
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data before and After save is different");

        }

        [Test]
        public void TestSerializeDrawingBigger8k_noAggregation()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("DrawingContinue.xls");

            InternalSheet isheet = HSSFTestHelper.GetSheetForTest(wb.GetSheetAt(0) as HSSFSheet);
            List<RecordBase> records = isheet.Records;

            HSSFWorkbook wb2 = HSSFTestDataSamples.WriteOutAndReadBack(wb);
            InternalSheet isheet2 = HSSFTestHelper.GetSheetForTest(wb2.GetSheetAt(0) as HSSFSheet);
            List<RecordBase> records2 = isheet2.Records;

            Assert.AreEqual(records.Count, records2.Count);
            for (int i = 0; i < records.Count; i++)
            {
                RecordBase r1 = records[(i)];
                RecordBase r2 = records2[(i)];
                Assert.IsTrue(r1.GetType() == r2.GetType());
                Assert.AreEqual(r1.RecordSize, r2.RecordSize);
                if (r1 is NPOI.HSSF.Record.Record)
                {
                    Assert.AreEqual(((NPOI.HSSF.Record.Record)r1).Sid, ((NPOI.HSSF.Record.Record)r2).Sid);
                    Assert.IsTrue(Arrays.Equals(((NPOI.HSSF.Record.Record)r1).Serialize(), ((NPOI.HSSF.Record.Record)r2).Serialize()));
                }
            }
        }
        [Test]
        public void TestSerializeDrawingWithComments()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("DrawingAndComments.xls");
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;
            InternalWorkbook iworkbook = HSSFTestHelper.GetWorkbookForTest(wb);
            InternalSheet isheet = HSSFTestHelper.GetSheetForTest(sh);

            List<RecordBase> records = isheet.Records;

            // the sheet's Drawing is not aggregated
            Assert.AreEqual(46, records.Count, "wrong size of sheet records stream");
            // the last record before the Drawing block
            Assert.IsTrue(records[(18)] is RowRecordsAggregate,
                "records.Get(18) is expected to be RowRecordsAggregate but was " + records[(18)].GetType().Name);

            // records to be aggregated
            List<RecordBase> dgRecords = records.GetRange(19, 39 - 19);
            foreach (RecordBase rb in dgRecords)
            {
                NPOI.HSSF.Record.Record r = (NPOI.HSSF.Record.Record)rb;
                short sid = r.Sid;
                // we expect that Drawing block consists of either
                // DrawingRecord or ContinueRecord or ObjRecord or TextObjectRecord
                Assert.IsTrue(
                        sid == DrawingRecord.sid ||
                                sid == ContinueRecord.sid ||
                                sid == ObjRecord.sid ||
                                sid == NoteRecord.sid ||
                                sid == TextObjectRecord.sid);
            }
            // collect Drawing records into a byte buffer.
            byte[] dgBytes = ToArray(dgRecords);

            // the first record After the Drawing block
            Assert.IsTrue(records[(39)] is WindowTwoRecord, "records.Get(39) is expected to be Window2");

            // aggregate Drawing records.
            // The subrange [19, 38] is expected to be Replaced with a EscherAggregate object
            DrawingManager2 drawingManager = iworkbook.FindDrawingGroup();
            int loc = isheet.AggregateDrawingRecords(drawingManager, false);
            EscherAggregate agg = (EscherAggregate)records[(loc)];

            Assert.AreEqual(27, records.Count, "wrong size of the aggregated sheet records stream");
            Assert.IsTrue(records[(18)] is RowRecordsAggregate,
                "records.Get(18) is expected to be RowRecordsAggregate but was " + records[(18)].GetType().Name);
            Assert.IsTrue(records[(19)] is EscherAggregate,
                "records.Get(19) is expected to be EscherAggregate but was " + records[(19)].GetType().Name);
            Assert.IsTrue(records[(20)] is WindowTwoRecord,
                "records.Get(20) is expected to be Window2 but was " + records[20].GetType().Name);

            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length, "different size of Drawing data before and After save");
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data before and After save is different");
        }

        [Test]
        public void TestFileWithPictures()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ContinueRecordProblem.xls");
            HSSFSheet sh = wb.GetSheetAt(0) as HSSFSheet;

            InternalWorkbook iworkbook = HSSFTestHelper.GetWorkbookForTest(wb);
            InternalSheet isheet = HSSFTestHelper.GetSheetForTest(sh);

            List<RecordBase> records = isheet.Records;

            // the sheet's Drawing is not aggregated
            Assert.AreEqual(315, records.Count, "wrong size of sheet records stream");
            // the last record before the Drawing block
            Assert.IsTrue(records[(21)] is RowRecordsAggregate,
                "records.Get(21) is expected to be RowRecordsAggregate but was " + records[(21)].GetType().Name);

            // records to be aggregated
            List<RecordBase> dgRecords = records.GetRange(22, 300 - 22);
            foreach (RecordBase rb in dgRecords)
            {
                NPOI.HSSF.Record.Record r = (NPOI.HSSF.Record.Record)rb;
                short sid = r.Sid;
                // we expect that Drawing block consists of either
                // DrawingRecord or ContinueRecord or ObjRecord or TextObjectRecord
                Assert.IsTrue(
                        sid == DrawingRecord.sid ||
                                sid == ContinueRecord.sid ||
                                sid == ObjRecord.sid ||
                                sid == TextObjectRecord.sid);
            }
            // collect Drawing records into a byte buffer.
            byte[] dgBytes = ToArray(dgRecords);

            // the first record After the Drawing block
            Assert.IsTrue(records[(300)] is WindowTwoRecord, "records.Get(300) is expected to be Window2");

            // aggregate Drawing records.
            // The subrange [19, 299] is expected to be Replaced with a EscherAggregate object
            DrawingManager2 drawingManager = iworkbook.FindDrawingGroup();
            int loc = isheet.AggregateDrawingRecords(drawingManager, false);
            EscherAggregate agg = (EscherAggregate)records[(loc)];

            Assert.AreEqual(38, records.Count, "wrong size of the aggregated sheet records stream");
            Assert.IsTrue(records[(21)] is RowRecordsAggregate,
                "records.Get(21) is expected to be RowRecordsAggregate but was " + records[(21)].GetType().Name);
            Assert.IsTrue(records[(22)] is EscherAggregate,
                "records.Get(22) is expected to be EscherAggregate but was " + records[22].GetType().Name);
            Assert.IsTrue(records[23] is WindowTwoRecord,
            "records.Get(23) is expected to be Window2 but was " + records[23].GetType().Name);

            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length, "different size of Drawing data before and After save");
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data before and After save is different");
        }
        [Test]
        public void TestUnhandledContinue()
        {
            String data =
                    "     EC 00 1C 08 0F 00 02 F0 66 27 00                " +
                            "     00 10 00 08 F0 08 00 00 00 06 00 00 00 13 04 00 " +
                            "     00 0F 00 03 F0 4E 27 00 00 0F 00 04 F0 28 00 00 " +
                            "     00 01 00 09 F0 10 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 00 00 00 00 00 00 00 00 02 00 0A F0 08 00 00 " +
                            "     00 00 04 00 00 05 00 00 00 0F 00 04 F0 CC 07 00 " +
                            "     00 A2 0C 0A F0 08 00 00 00 0F 04 00 00 00 0A 00 " +
                            "     00 93 00 0B F0 4A 00 00 00 7F 00 00 00 04 00 80 " +
                            "     00 C0 D5 9A 02 85 00 02 00 00 00 8B 00 02 00 00 " +
                            "     00 BF 00 1A 00 1F 00 BF 01 01 00 11 00 FF 01 00 " +
                            "     00 08 00 80 C3 14 00 00 00 BF 03 00 00 02 00 54 " +
                            "     00 65 00 78 00 74 00 42 00 6F 00 78 00 20 00 31 " +
                            "     00 00 00 13 00 22 F1 38 07 00 00 A9 C3 32 07 00 " +
                            "     00 50 4B 03 04 14 00 06 00 08 00 00 00 21 00 F0 " +
                            "     F7 8A BB FD 00 00 00 E2 01 00 00 13 00 00 00 5B " +
                            "     43 6F 6E 74 65 6E 74 5F 54 79 70 65 73 5D 2E 78 " +
                            "     6D 6C 94 91 CD 4A C4 30 10 C7 EF 82 EF 10 E6 2A " +
                            "     6D AA 07 11 69 BA 07 AB 47 15 5D 1F 60 48 A6 6D " +
                            "     D8 36 09 99 58 77 DF DE 74 3F 2E E2 0A 1E 67 E6 " +
                            "     FF F1 23 A9 57 DB 69 14 33 45 B6 DE 29 B8 2E 2B " +
                            "     10 E4 B4 37 D6 F5 0A 3E D6 4F C5 1D 08 4E E8 0C " +
                            "     8E DE 91 82 1D 31 AC 9A CB 8B 7A BD 0B C4 22 BB " +
                            "     1D 2B 18 52 0A F7 52 B2 1E 68 42 2E 7D 20 97 2F " +
                            "     9D 8F 13 A6 3C C6 5E 06 D4 1B EC 49 DE 54 D5 AD " +
                            "     D4 DE 25 72 A9 48 4B 06 34 75 4B 1D 7E 8E 49 3C " +
                            "     6E F3 FA 40 12 69 64 10 0F 07 E1 D2 A5 00 43 18 " +
                            "     AD C6 94 49 E5 EC CC 8F 96 E2 D8 50 66 E7 5E C3 " +
                            "     83 0D 7C 95 31 40 FE DA B0 5C CE 17 1C 7D 2F F9 " +
                            "     69 A2 35 24 5E 31 A6 67 9C 32 86 34 91 25 0F 18 " +
                            "     28 6B CA BF 53 16 CC 89 0B DF 75 56 53 D9 46 7E " +
                            "     5F 7C 27 A8 73 E1 C6 7F B9 48 F3 7F B3 DB 6C 7B " +
                            "     A3 F9 94 2E F7 3F D4 7C 03 00 00 FF FF 03 00 50 " +
                            "     4B 03 04 14 00 06 00 08 00 00 00 21 00 31 DD 5F " +
                            "     61 D2 00 00 00 8F 01 00 00 0B 00 00 00 5F 72 65 " +
                            "     6C 73 2F 2E 72 65 6C 73 A4 90 C1 6A C3 30 0C 86 " +
                            "     EF 83 BD 83 D1 BD 71 DA 43 19 A3 4E 6F 85 5E 4B " +
                            "     07 BB 0A 5B 49 4C 63 CB 58 26 6D DF BE A6 30 58 " +
                            "     46 6F 3B EA 17 FA 3E F1 EF F6 B7 30 A9 99 B2 78 " +
                            "     8E 06 D6 4D 0B 8A A2 65 E7 E3 60 E0 EB 7C 58 7D " +
                            "     80 92 82 D1 E1 C4 91 0C DC 49 60 DF BD BF ED 4E " +
                            "     34 61 A9 47 32 FA 24 AA 52 A2 18 18 4B 49 9F 5A " +
                            "     8B 1D 29 A0 34 9C 28 D6 4D CF 39 60 A9 63 1E 74 " +
                            "     42 7B C1 81 F4 A6 6D B7 3A FF 66 40 B7 60 AA A3 " +
                            "     33 90 8F 6E 03 EA 7C 4F D5 FC 87 1D BC CD 2C DC " +
                            "     97 C6 72 D0 DC F7 DE BE A2 6A C7 D7 78 A2 B9 52 " +
                            "     30 0F 54 0C B8 2C CF 30 D3 DC D4 E7 40 BF F6 AE " +
                            "     FF E9 95 11 13 7D 57 FE 42 FC 4C AB F5 C7 AC 17 " +
                            "     35 76 0F 00 00 00 FF FF 03 00 50 4B 03 04 14 00 " +
                            "     06 00 08 00 00 00 21 00 AE B1 05 56 77 02 00 00 " +
                            "     1B 07 00 00 10 00 00 00 64 72 73 2F 73 68 61 70 " +
                            "     65 78 6D 6C 2E 78 6D 6C BC 55 4B 6F DB 30 0C BE " +
                            "     0F D8 7F 10 74 4F FD 58 9D 87 11 BB E8 52 74 97 " +
                            "     61 0D 92 E6 07 A8 B2 9C 18 93 25 43 52 13 A7 BF " +
                            "     7E A4 E4 A4 5D 0F 3B 34 43 2F 0A 43 4A FC C8 8F " +
                            "     0F CF 6F FA 56 92 BD 30 B6 D1 AA A0 C9 55 4C 89 " +
                            "     50 5C 57 8D DA 16 74 F3 78 3F 9A 52 62 1D 53 15 " +
                            "     93 5A 89 82 1E 85 A5 37 E5 D7 2F F3 BE 32 39 53 " +
                            "     7C A7 0D 01 17 CA E6 A0 28 E8 CE B9 2E 8F 22 CB " +
                            "     77 A2 65 F6 4A 77 42 81 B5 D6 A6 65 0E FE 9A 6D " +
                            "     54 19 76 00 E7 AD 8C D2 38 1E 47 B6 33 82 55 76 " +
                            "     27 84 BB 0B 16 5A 7A DF 80 B6 10 52 DE 7A 88 A0 " +
                            "     AA 8D 6E 83 C4 B5 2C 93 79 84 31 A0 E8 1F 80 F0 " +
                            "     50 D7 65 1A CF B2 2C 3E DB 50 E5 CD 46 1F CA 6F " +
                            "     41 8D E2 49 87 F6 64 1A CF 26 D9 D9 E6 9F 78 DF " +
                            "     AF 80 A2 77 84 F7 05 CD D2 2C 8D 81 12 7E 2C 68 " +
                            "     3A BE CE C6 31 8D 82 2F DB 91 96 71 A3 0B 4A 89 " +
                            "     83 EB B2 51 BF 41 0E 46 B5 5F 77 4B 13 64 FE 6B " +
                            "     BF 34 A4 A9 C0 01 25 8A B5 40 EA 23 DC FF AE 7B " +
                            "     92 9C 9C C1 1D 7C 40 5C 0F 6A 28 0B EA 7D 44 6F " +
                            "     1D 59 EF 92 E5 7D 6D DA A1 06 EC 03 15 68 59 A3 " +
                            "     20 4C 96 EB BA 26 00 36 4D 66 49 06 6D 00 19 4E " +
                            "     B2 F4 7A 92 21 38 CB FF CD 40 14 C2 C0 8B 9D B1 " +
                            "     EE 87 D0 17 87 44 D0 51 41 8D E0 CE 87 C7 F6 3F " +
                            "     AD 43 1E 5E 21 10 4E E9 FB 46 CA 4B F3 3F F1 1B " +
                            "     38 C5 BE B2 EE 28 05 02 48 B5 12 40 8C EF F1 0F " +
                            "     F3 0B F5 06 6A 63 9F 88 E5 66 FB B4 90 86 C0 BC " +
                            "     00 CF 30 67 70 3E E1 19 B2 F3 80 88 5C 43 62 9F " +
                            "     8C 3D 40 22 BA A8 6B A0 FE 93 F1 CF A0 3E 7F AD " +
                            "     FE 1F 7E DB 28 6D 06 FE 61 39 09 2C C0 9E C9 82 " +
                            "     BA 3E 8C 17 F0 1D F0 86 51 1B 1A 00 7B 01 C7 B0 " +
                            "     3A 62 48 4F F0 0B 73 79 69 37 C0 C2 75 0F 70 D4 " +
                            "     52 1F 0A CA 65 D3 51 02 9B F4 E5 BD EE 60 58 57 " +
                            "     50 05 BB 90 12 E3 E4 42 43 BC D0 2C 61 EF 42 E8 " +
                            "     21 9F EE F6 D9 C1 14 0C C3 11 42 C4 60 A5 75 6B " +
                            "     EC E2 4B C3 F5 F3 DF 5D EA 05 23 32 FE 00 02 25 " +
                            "     C3 EF 8B 50 A3 CD 1A BE 2F 2F B0 E2 92 D8 F7 3F " +
                            "     CB 5D 89 EB 73 04 EB 1D 65 3C FD 2B A1 AA 25 33 " +
                            "     6C 75 7E 6C 9E 47 AB CD DF 8F 71 35 0C D5 3B 95 " +
                            "     CC 6F 4D 0B 5A FF AD 90 8D 50 EE 8E 39 76 9A F7 " +
                            "     77 5F 19 7F 3B B0 5B FE 01 00 00 FF FF 03 00 50 " +
                            "     4B 03 04 14 00 06 00 08 00 00 00 21 00 44 8B 69 " +
                            "     1E 2C 01 00 00 AB 01 00 00 0F 00 00 00 64 72 73 " +
                            "     2F 64 6F 77 6E 72 65 76 2E 78 6D 6C 5C 50 D1 4E " +
                            "     C2 30 14 7D 37 F1 1F 9A 6B E2 9B 74 0C 36 07 52 " +
                            "     C8 62 42 C4 04 D1 21 89 AF 65 6D D9 E2 DA CE B6 " +
                            "     C0 F0 EB ED 46 0C C6 C7 73 EE 39 E7 DE 73 27 B3 " +
                            "     46 56 E8 C0 8D 2D B5 22 D0 EF 05 80 B8 CA 35 2B " +
                            "     D5 8E C0 E6 7D 7E 97 00 B2 8E 2A 46 2B AD 38 81 " +
                            "     13 B7 30 9B 5E 5F 4D E8 98 E9 A3 CA F8 61 ED 76 " +
                            "     C8 87 28 3B A6 04 0A E7 EA 31 C6 36 2F B8 A4 B6 " +
                            "     A7 6B AE FC 4C 68 23 A9 F3 D0 EC 30 33 F4 E8 C3 " +
                            "     65 85 C3 20 88 B1 A4 A5 F2 1B 0A 5A F3 C7 82 E7 " +
                            "     9F EB BD 24 B0 64 E2 E5 63 58 CE 9F A9 F8 4A 17 " +
                            "     C3 2C 89 0E AF 2C 22 E4 F6 A6 49 1F 00 39 DE B8 " +
                            "     8B B8 DE AB 55 B8 5C 6E DF 36 E9 AF A0 8B 5B 30 " +
                            "     02 21 20 F1 74 DA 9A 92 65 D4 3A 6E 08 F8 7E BE " +
                            "     AD 6F 0A 53 5F A1 A9 52 95 17 DA 20 91 71 5B 7E " +
                            "     FB 7E 67 5E 18 2D 91 D1 47 02 03 40 B9 AE 5A BE " +
                            "     C5 2B 21 2C 77 1E 25 C1 E8 3E EA 46 BF 54 18 8C " +
                            "     A2 28 00 DC C6 3A 7D 36 9F 15 DD 19 7F CC 71 1C " +
                            "     FF F3 F6 C3 41 12 46 AD 17 5F 6E EA C0 E5 C7 D3 " +
                            "     1F 00 00 00 FF FF 03 00 50 4B 01 02 2D 00 14 00 " +
                            "     06 00 08 00 00 00 21 00 F0 F7 8A BB FD 00 00 00 " +
                            "     E2 01 00 00 13 00 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 00 00 00 00 00 5B 43 6F 6E 74 65 6E 74 5F 54 " +
                            "     79 70 65 73 5D 2E 78 6D 6C 50 4B 01 02 2D 00 14 " +
                            "     00 06 00 08 00 00 00 21 00 31 DD 5F 61 D2 00 00 " +
                            "     00 8F 01 00 00 0B 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 00 00 2E 01 00 00 5F 72 65 6C 73 2F 2E 72 65 " +
                            "     6C 73 50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 " +
                            "     21 00 AE B1 05 56 77 02 00 00 1B 07 00 00 10 00 " +
                            "     00 00 00 00 00 00 00 00 00 00 00 00 29 02 00 00 " +
                            "     64 72 73 2F 73 68 61 70 65 78 6D 6C 2E 78 6D 6C " +
                            "     50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 21 00 " +
                            "     44 8B 69 1E 2C 01 00 00 AB 01 00 00 0F 00 00 00 " +
                            "     00 00 00 00 00 00 00 00 00 00 CE 04 00 00 64 72 " +
                            "     73 2F 64 6F 77 6E 72 65 76 2E 78 6D 6C 50 4B 05 " +
                            "     06 00 00 00 00 04 00 04 00 F5 00 00 00 27 06 00 " +
                            "     00 00 00 00 00 10 F0 12 00 00 00 02 00 01 00 60 " +
                            "     01 03 00 F3 00 02 00 D0 00 05 00 5A 00 00 00 11 " +
                            "     F0 00 00 00 00 5D 00 1A 00 15 00 12 00 06 00 0F " +
                            "     00 11 60 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 00 00 EC 00 08 00 00 00 0D F0 00 00 00 00 B6 " +
                            "     01 12 00 12 02 00 00 00 00 00 00 00 00 06 00 10 " +
                            "     00 00 00 00 00 3C 00 07 00 00 74 65 78 74 2D 31 " +
                            "     3C 00 10 00 00 00 16 00 00 00 00 00 06 00 00 00 " +
                            "     00 00 00 00 EC 00 CB 07 0F 00 04 F0 CB 07 00 00 " +
                            "     A2 0C 0A F0 08 00 00 00 10 04 00 00 00 0A 00 00 " +
                            "     93 00 0B F0 4A 00 00 00 7F 00 00 00 04 00 80 00 " +
                            "     00 65 53 02 85 00 02 00 00 00 8B 00 02 00 00 00 " +
                            "     BF 00 1A 00 1F 00 BF 01 01 00 11 00 FF 01 00 00 " +
                            "     08 00 80 C3 14 00 00 00 BF 03 00 00 02 00 54 00 " +
                            "     65 00 78 00 74 00 42 00 6F 00 78 00 20 00 32 00 " +
                            "     00 00 13 00 22 F1 37 07 00 00 A9 C3 31 07 00 00 " +
                            "     50 4B 03 04 14 00 06 00 08 00 00 00 21 00 F0 F7 " +
                            "     8A BB FD 00 00 00 E2 01 00 00 13 00 00 00 5B 43 " +
                            "     6F 6E 74 65 6E 74 5F 54 79 70 65 73 5D 2E 78 6D " +
                            "     6C 94 91 CD 4A C4 30 10 C7 EF 82 EF 10 E6 2A 6D " +
                            "     AA 07 11 69 BA 07 AB 47 15 5D 1F 60 48 A6 6D D8 " +
                            "     36 09 99 58 77 DF DE 74 3F 2E E2 0A 1E 67 E6 FF " +
                            "     F1 23 A9 57 DB 69 14 33 45 B6 DE 29 B8 2E 2B 10 " +
                            "     E4 B4 37 D6 F5 0A 3E D6 4F C5 1D 08 4E E8 0C 8E " +
                            "     DE 91 82 1D 31 AC 9A CB 8B 7A BD 0B C4 22 BB 1D " +
                            "     2B 18 52 0A F7 52 B2 1E 68 42 2E 7D 20 97 2F 9D " +
                            "     8F 13 A6 3C C6 5E 06 D4 1B EC 49 DE 54 D5 AD D4 " +
                            "     DE 25 72 A9 48 4B 06 34 75 4B 1D 7E 8E 49 3C 6E " +
                            "     F3 FA 40 12 69 64 10 0F 07 E1 D2 A5 00 43 18 AD " +
                            "     C6 94 49 E5 EC CC 8F 96 E2 D8 50 66 E7 5E C3 83 " +
                            "     0D 7C 95 31 40 FE DA B0 5C CE 17 1C 7D 2F F9 69 " +
                            "     A2 35 24 5E 31 A6 67 9C 32 86 34 91 25 0F 18 28 " +
                            "     6B CA BF 53 16 CC 89 0B DF 75 56 53 D9 46 7E 5F " +
                            "     7C 27 A8 73 E1 C6 7F B9 48 F3 7F B3 DB 6C 7B A3 " +
                            "     F9 94 2E F7 3F D4 7C 03 00 00 FF FF 03 00 50 4B " +
                            "     03 04 14 00 06 00 08 00 00 00 21 00 31 DD 5F 61 " +
                            "     D2 00 00 00 8F 01 00 00 0B 00 00 00 5F 72 65 6C " +
                            "     73 2F 2E 72 65 6C 73 A4 90 C1 6A C3 30 0C 86 EF " +
                            "     83 BD 83 D1 BD 71 DA 43 19 A3 4E 6F 85 5E 4B 07 " +
                            "     BB 0A 5B 49 4C 63 CB 58 26 6D DF BE A6 30 58 46 " +
                            "     6F 3B EA 17 FA 3E F1 EF F6 B7 30 A9 99 B2 78 8E " +
                            "     06 D6 4D 0B 8A A2 65 E7 E3 60 E0 EB 7C 58 7D 80 " +
                            "     92 82 D1 E1 C4 91 0C DC 49 60 DF BD BF ED 4E 34 " +
                            "     61 A9 47 32 FA 24 AA 52 A2 18 18 4B 49 9F 5A 8B " +
                            "     1D 29 A0 34 9C 28 D6 4D CF 39 60 A9 63 1E 74 42 " +
                            "     7B C1 81 F4 A6 6D B7 3A FF 66 40 B7 60 AA A3 33 " +
                            "     90 8F 6E 03 EA 7C 4F D5 FC 87 1D BC CD 2C DC 97 " +
                            "     C6 72 D0 DC F7 DE BE A2 6A C7 D7 78 A2 B9 52 30 " +
                            "     0F 54 0C B8 2C CF 30 D3 DC D4 E7 40 BF F6 AE FF " +
                            "     E9 95 11 13 7D 57 FE 42 FC 4C AB F5 C7 AC 17 35 " +
                            "     76 0F 00 00 00 FF FF 03 00 50 4B 03 04 14 00 06 " +
                            "     00 08 00 00 00 21 00 99 C9 E2 87 75 02 00 00 1B " +
                            "     07 00 00 10 00 00 00 64 72 73 2F 73 68 61 70 65 " +
                            "     78 6D 6C 2E 78 6D 6C BC 55 CD 6E E2 30 10 BE AF " +
                            "     D4 77 B0 7C A7 F9 E9 82 20 C2 A9 5A AA EE 65 B5 " +
                            "     45 50 1E C0 24 0E 44 EB D8 91 ED 42 E8 D3 EF 8C " +
                            "     9D D0 6E 0F 7B 28 AB 5E 1C 67 C6 33 DF F8 9B 1F " +
                            "     CF 6F BB 46 92 83 30 B6 D6 8A D1 E4 3A A6 44 A8 " +
                            "     42 97 B5 DA 31 BA 79 7E 1C 4D 29 B1 8E AB 92 4B " +
                            "     AD 04 A3 27 61 E9 6D 7E F5 6D DE 95 26 E3 AA D8 " +
                            "     6B 43 C0 85 B2 19 08 18 DD 3B D7 66 51 64 8B BD " +
                            "     68 B8 BD D6 AD 50 A0 AD B4 69 B8 83 5F B3 8B 4A " +
                            "     C3 8F E0 BC 91 51 1A C7 93 C8 B6 46 F0 D2 EE 85 " +
                            "     70 0F 41 43 73 EF 1B D0 16 42 CA 3B 0F 11 44 95 " +
                            "     D1 4D D8 15 5A E6 C9 3C C2 18 70 EB 0D 60 F3 54 " +
                            "     55 79 1A CF C6 E3 F8 AC 43 91 57 1B 7D CC 27 41 " +
                            "     8C DB 41 86 FA D9 38 1D 2C 40 E5 2D BC EB 37 3C " +
                            "     D1 39 52 74 8C C2 C1 34 06 46 8A 13 A3 E9 E4 FB " +
                            "     78 12 D3 28 B8 B2 2D 69 78 61 34 A3 94 38 38 2E " +
                            "     6B F5 1B F6 41 A9 0E EB 76 69 C2 BE F8 75 58 1A " +
                            "     52 97 8C DE 50 A2 78 03 9C 3E C3 F9 7B DD 91 74 " +
                            "     70 06 67 D0 80 B8 0E C4 90 15 94 FB 88 DE 3B B2 " +
                            "     DE 25 CF BA CA 34 7D 0A F8 27 12 D0 F0 5A 41 98 " +
                            "     3C D3 55 45 00 6C 9A CC 92 31 54 01 DC 30 49 6F " +
                            "     A6 C0 0C A2 F3 EC DF 14 44 21 0E 3C D8 1A EB 7E " +
                            "     08 7D 71 4C 04 1D 31 6A 44 E1 7C 7C FC F0 D3 3A " +
                            "     24 E2 0D 02 E1 94 7E AC A5 BC 94 80 81 E0 40 2A " +
                            "     D6 95 75 27 29 10 40 AA 95 00 66 7C 8D 7F 9A 60 " +
                            "     48 38 70 1B FB 8B D8 C2 EC B6 0B 69 08 F4 0B 10 " +
                            "     0D 7D 06 EB 16 D7 70 3B 0F 88 C8 15 5C EC 8B B1 " +
                            "     7B 48 44 17 55 05 D4 7F 31 FE 19 D4 DF 5F AB FF " +
                            "     87 DF D4 4A 9B 9E 7F 18 4E 02 13 70 E0 92 51 D7 " +
                            "     85 FE 02 BE 03 5E DF 6B 7D 01 60 2D 60 1F 96 27 " +
                            "     0C 69 0B 5F 68 CC 4B AB 01 06 AE 7B 82 A5 92 FA " +
                            "     C8 68 21 EB 96 12 98 A4 AF 1F 65 47 C3 5B 46 15 " +
                            "     CC 42 4A 8C 93 0B 0D F1 42 B1 84 B9 0B A1 87 FB " +
                            "     B4 77 2F 0E BA A0 6F 8E 10 22 06 2B AD 5B 63 15 " +
                            "     5F 1A AE EF FF F6 52 2F 18 91 F1 0B 10 28 39 BE " +
                            "     2F 42 8D 36 6B 78 5F 5E 61 D6 24 71 3F 68 5C 8E " +
                            "     F3 73 94 62 A3 3B DF EE DE 4A A8 72 C9 0D 5F 9D " +
                            "     8D CD CB 68 B5 F9 DB 18 47 43 9F BD 21 65 7E 6C " +
                            "     5A 90 FA B7 42 D6 42 B9 07 EE F8 D0 EF 1F 5E 19 " +
                            "     7F 3A B0 9B FF 01 00 00 FF FF 03 00 50 4B 03 04 " +
                            "     14 00 06 00 08 00 00 00 21 00 56 84 64 08 2D 01 " +
                            "     00 00 AB 01 00 00 0F 00 00 00 64 72 73 2F 64 6F " +
                            "     77 6E 72 65 76 2E 78 6D 6C 5C 90 5F 4F C2 30 14 " +
                            "     C5 DF 4D FC 0E CD 35 F1 4D B6 15 C6 3F 29 64 31 " +
                            "     21 62 82 E2 90 C4 D7 B2 B6 6C BA B6 B3 AD 30 F8 " +
                            "     F4 76 20 21 FA 78 CE BD BF D3 7B 3A 9A D4 B2 44 " +
                            "     5B 6E 6C A1 15 81 A8 15 02 E2 2A D3 AC 50 1B 02 " +
                            "     AB B7 E9 5D 1F 90 75 54 31 5A 6A C5 09 EC B9 85 " +
                            "     C9 F8 FA 6A 44 87 4C EF 54 CA B7 4B B7 41 3E 44 " +
                            "     D9 21 25 90 3B 57 0D 83 C0 66 39 97 D4 B6 74 C5 " +
                            "     95 9F 09 6D 24 75 5E 9A 4D C0 0C DD F9 70 59 06 " +
                            "     38 0C BB 81 A4 85 F2 2F E4 B4 E2 0F 39 CF 3E 97 " +
                            "     DF 92 C0 9C 89 E7 F7 4E 31 7D A2 E2 2B 99 75 D2 " +
                            "     7E BC 5D B0 98 90 DB 9B 3A B9 07 E4 78 ED 2E CB " +
                            "     55 FE B1 C0 F3 F9 FA 75 95 9C 17 8E 71 33 46 A0 " +
                            "     0D 48 3C EE D7 A6 60 29 B5 8E 1B 02 BE 9F 6F EB " +
                            "     9B C2 D8 57 A8 CB 44 65 B9 36 48 A4 DC 16 07 DF " +
                            "     EF E4 0B A3 25 32 7A 47 A0 0B 28 D3 65 E3 37 FA " +
                            "     45 08 CB 1D 81 41 8C 63 9F E4 27 67 07 87 83 D8 " +
                            "     5B 41 93 EA F4 89 ED FD B2 F8 0F 1B F5 A2 CE 3F " +
                            "     38 C2 ED 3E 8E 1B 38 B8 DC 74 14 97 3F 1E FF 00 " +
                            "     00 00 FF FF 03 00 50 4B 01 02 2D 00 14 00 06 00 " +
                            "     08 00 00 00 21 00 F0 F7 8A BB FD 00 00 00 E2 01 " +
                            "     00 00 13 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 00 00 00 5B 43 6F 6E 74 65 6E 74 5F 54 79 70 " +
                            "     65 73 5D 2E 78 6D 6C 50 4B 01 02 2D 00 14 00 06 " +
                            "     00 08 00 00 00 21 00 31 DD 5F 61 D2 00 00 00 8F " +
                            "     01 00 00 0B 00 00 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 2E 01 00 00 5F 72 65 6C 73 2F 2E 72 65 6C 73 " +
                            "     50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 21 00 " +
                            "     99 C9 E2 87 75 02 00 00 1B 07 00 00 10 00 00 00 " +
                            "     00 00 00 00 00 00 00 00 00 00 29 02 00 00 64 72 " +
                            "     73 2F 73 68 61 70 65 78 6D 6C 2E 78 6D 6C 50 4B " +
                            "     01 02 2D 00 14 00 06 00 08 00 00 00 21 00 56 84 " +
                            "     64 08 2D 01 00 00 AB 01 00 00 0F 00 00 00 00 00 " +
                            "     00 00 00 00 00 00 00 00 CC 04 00 00 64 72 73 2F " +
                            "     64 6F 77 6E 72 65 76 2E 78 6D 6C 50 4B 05 06 00 " +
                            "     00 00 00 04 00 04 00 F5 00 00 00 26 06 00 00 00 " +
                            "     00 00 00 10 F0 12 00 00 00 02 00 01 00 60 01 06 " +
                            "     00 80 00 02 00 D0 00 07 00 E6 00 00 00 11 F0 00 " +
                            "     00 00 00 5D 00 1A 00 15 00 12 00 06 00 10 00 11 " +
                            "     60 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 EC 00 08 00 00 00 0D F0 00 00 00 00 B6 01 12 " +
                            "     00 12 02 00 00 00 00 00 00 00 00 06 00 10 00 00 " +
                            "     00 00 00 3C 00 07 00 00 74 65 78 74 2D 32 3C 00 " +
                            "     10 00 00 00 16 00 00 00 00 00 06 00 00 00 00 00 " +
                            "     00 00 EC 00 CB 07 0F 00 04 F0 CB 07 00 00 A2 0C " +
                            "     0A F0 08 00 00 00 11 04 00 00 00 0A 00 00 93 00 " +
                            "     0B F0 4A 00 00 00 7F 00 00 00 04 00 80 00 00 D1 " +
                            "     53 02 85 00 02 00 00 00 8B 00 02 00 00 00 BF 00 " +
                            "     1A 00 1F 00 BF 01 01 00 11 00 FF 01 00 00 08 00 " +
                            "     80 C3 14 00 00 00 BF 03 00 00 02 00 54 00 65 00 " +
                            "     78 00 74 00 42 00 6F 00 78 00 20 00 33 00 00 00 " +
                            "     13 00 22 F1 37 07 00 00 A9 C3 31 07 00 00 50 4B " +
                            "     03 04 14 00 06 00 08 00 00 00 21 00 F0 F7 8A BB " +
                            "     FD 00 00 00 E2 01 00 00 13 00 00 00 5B 43 6F 6E " +
                            "     74 65 6E 74 5F 54 79 70 65 73 5D 2E 78 6D 6C 94 " +
                            "     91 CD 4A C4 30 10 C7 EF 82 EF 10 E6 2A 6D AA 07 " +
                            "     11 69 BA 07 AB 47 15 5D 1F 60 48 A6 6D D8 36 09 " +
                            "     99 58 77 DF DE 74 3F 2E E2 0A 1E 67 E6 FF F1 23 " +
                            "     A9 57 DB 69 14 33 45 B6 DE 29 B8 2E 2B 10 E4 B4 " +
                            "     37 D6 F5 0A 3E D6 4F C5 1D 08 4E E8 0C 8E DE 91 " +
                            "     82 1D 31 AC 9A CB 8B 7A BD 0B C4 22 BB 1D 2B 18 " +
                            "     52 0A F7 52 B2 1E 68 42 2E 7D 20 97 2F 9D 8F 13 " +
                            "     A6 3C C6 5E 06 D4 1B EC 49 DE 54 D5 AD D4 DE 25 " +
                            "     72 A9 48 4B 06 34 75 4B 1D 7E 8E 49 3C 6E F3 FA " +
                            "     40 12 69 64 10 0F 07 E1 D2 A5 00 43 18 AD C6 94 " +
                            "     49 E5 EC CC 8F 96 E2 D8 50 66 E7 5E C3 83 0D 7C " +
                            "     95 31 40 FE DA B0 5C CE 17 1C 7D 2F F9 69 A2 35 " +
                            "     24 5E 31 A6 67 9C 32 86 34 91 25 0F 18 28 6B CA " +
                            "     BF 53 16 CC 89 0B DF 75 56 53 D9 46 7E 5F 7C 27 " +
                            "     A8 73 E1 C6 7F B9 48 F3 7F B3 DB 6C 7B A3 F9 94 " +
                            "     2E F7 3F D4 7C 03 00 00 FF FF 03 00 50 4B 03 04 " +
                            "     14 00 06 00 08 00 00 00 21 00 31 DD 5F 61 D2 00 " +
                            "     00 00 8F 01 00 00 0B 00 00 00 5F 72 65 6C 73 2F " +
                            "     2E 72 65 6C 73 A4 90 C1 6A C3 30 0C 86 EF 83 BD " +
                            "     83 D1 BD 71 DA 43 19 A3 4E 6F 85 5E 4B 07 BB 0A " +
                            "     5B 49 4C 63 CB 58 26 6D DF BE A6 30 58 46 6F 3B " +
                            "     EA 17 FA 3E F1 EF F6 B7 30 A9 99 B2 78 8E 06 D6 " +
                            "     4D 0B 8A A2 65 E7 E3 60 E0 EB 7C 58 7D 80 92 82 " +
                            "     D1 E1 C4 91 0C DC 49 60 DF BD BF ED 4E 34 61 A9 " +
                            "     47 32 FA 24 AA 52 A2 18 18 4B 49 9F 5A 8B 1D 29 " +
                            "     A0 34 9C 28 D6 4D CF 39 60 A9 63 1E 74 42 7B C1 " +
                            "     81 F4 A6 6D B7 3A FF 66 40 B7 60 AA A3 33 90 8F " +
                            "     6E 03 EA 7C 4F D5 FC 87 1D BC CD 2C DC 97 C6 72 " +
                            "     D0 DC F7 DE BE A2 6A C7 D7 78 A2 B9 52 30 0F 54 " +
                            "     0C B8 2C CF 30 D3 DC D4 E7 40 BF F6 AE FF E9 95 " +
                            "     11 13 7D 57 FE 42 FC 4C AB F5 C7 AC 17 35 76 0F " +
                            "     00 00 00 FF FF 03 00 50 4B 03 04 14 00 06 00 08 " +
                            "     00 00 00 21 00 49 0D 41 2E 77 02 00 00 1A 07 00 " +
                            "     00 10 00 00 00 64 72 73 2F 73 68 61 70 65 78 6D " +
                            "     6C 2E 78 6D 6C BC 55 4D 73 DB 20 10 BD 77 A6 FF " +
                            "     81 E1 EE E8 23 96 1B 6B 8C 32 A9 33 E9 A5 D3 78 " +
                            "     EC F8 07 60 09 D9 9A 22 D0 00 B1 E5 FC FA EE 82 " +
                            "     EC A4 39 F4 10 77 72 41 68 17 F6 3D 1E BB CB EC " +
                            "     B6 6F 25 D9 0B 63 1B AD 18 4D AE 62 4A 84 2A 75 " +
                            "     D5 A8 2D A3 EB A7 87 D1 0D 25 D6 71 55 71 A9 95 " +
                            "     60 F4 28 2C BD 2D BE 7E 99 F5 95 C9 B9 2A 77 DA " +
                            "     10 08 A1 6C 0E 06 46 77 CE 75 79 14 D9 72 27 5A " +
                            "     6E AF 74 27 14 78 6B 6D 5A EE E0 D7 6C A3 CA F0 " +
                            "     03 04 6F 65 94 C6 F1 24 B2 9D 11 BC B2 3B 21 DC " +
                            "     7D F0 D0 C2 C7 06 B4 B9 90 F2 CE 43 04 53 6D 74 " +
                            "     1B 66 A5 96 45 32 8B 90 03 4E FD 06 98 3C D6 75 " +
                            "     91 C6 D3 2C 8B CF 3E 34 79 B7 D1 87 62 1A CC 38 " +
                            "     3D D9 D0 3F CD D2 EC EC F1 1B 7C E4 57 38 D1 3B " +
                            "     52 F6 8C C2 BA 34 06 41 CA 23 A3 E9 64 9C 4D 62 " +
                            "     1A 85 48 B6 23 2D 2F 8D 66 94 12 07 CB 65 A3 7E " +
                            "     C3 3C 38 D5 7E D5 2D 4C 98 97 BF F6 0B 43 9A 8A " +
                            "     D1 31 25 8A B7 20 E9 13 AC FF AE 7B 72 7D 0A 06 " +
                            "     6B 70 03 71 3D 98 E1 52 D0 EE 19 BD 0D 64 7D 48 " +
                            "     9E F7 B5 69 87 1B E0 1F D0 BF E5 8D 02 9A 3C D7 " +
                            "     75 4D 00 EC 26 99 26 19 24 01 9C 30 F9 96 8E E3 " +
                            "     34 43 74 9E FF 5B 82 28 F0 C0 85 9D B1 EE 87 D0 " +
                            "     17 73 22 18 88 51 23 4A E7 F9 F1 FD 4F EB 50 88 " +
                            "     57 08 84 53 FA A1 91 F2 52 01 4E 02 07 51 31 AD " +
                            "     AC 3B 4A 81 00 52 2D 05 28 E3 53 FC C3 02 C3 85 " +
                            "     83 B6 B1 3F 88 2D CD 76 33 97 86 40 B9 80 D0 50 " +
                            "     66 30 6E 70 0C A7 F3 80 88 5C C3 C1 3E 19 7B 80 " +
                            "     44 74 51 D7 20 FD 27 E3 9F 41 FD F9 B5 FA 7F F8 " +
                            "     6D A3 B4 19 F4 87 DE 24 F0 02 F6 5C 32 EA FA 50 " +
                            "     5F A0 77 C0 1B 6A 6D 48 00 CC 05 AC C3 EA 88 94 " +
                            "     36 F0 85 C2 BC 34 1B A0 DF BA 47 18 6A A9 0F 8C " +
                            "     96 B2 E9 28 81 46 FA F2 DE 76 30 BC 63 54 41 2B " +
                            "     A4 C4 38 39 D7 C0 17 92 25 B4 5D A0 1E CE D3 DD " +
                            "     3D 3B A8 82 A1 38 02 45 24 2B AD 5B 61 16 5F 4A " +
                            "     D7 D7 7F 77 69 14 64 64 FC 00 02 4A 8E CF 8B 50 " +
                            "     A3 F5 0A 9E 97 17 E8 35 49 EC F3 9F E7 AE C0 FE " +
                            "     39 BA C6 42 77 BE DC FD 2E A1 AA 05 37 7C 79 DE " +
                            "     6C 9E 47 CB F5 DF 9B B1 35 0C B7 77 BA 32 DF 36 " +
                            "     2D 58 FD 53 21 1B A1 DC 3D 77 FC 54 EF EF 1E 19 " +
                            "     BF 3A A8 5B FC 01 00 00 FF FF 03 00 50 4B 03 04 " +
                            "     14 00 06 00 08 00 00 00 21 00 79 45 76 AD 2B 01 " +
                            "     00 00 AA 01 00 00 0F 00 00 00 64 72 73 2F 64 6F " +
                            "     77 6E 72 65 76 2E 78 6D 6C 5C 90 51 4F C2 30 14 " +
                            "     85 DF 4D FC 0F CD 35 F1 4D 3A 26 D5 81 14 B2 98 " +
                            "     10 31 99 E2 90 C4 D7 B2 B6 6C 71 6D 67 5B 61 F8 " +
                            "     EB 2D 23 84 E8 E3 39 F7 7E E7 F6 74 3C 6D 55 8D " +
                            "     B6 C2 BA CA 68 0A FD 5E 04 48 E8 C2 F0 4A 6F 28 " +
                            "     AC DE 67 37 09 20 E7 99 E6 AC 36 5A 50 D8 0B 07 " +
                            "     D3 C9 E5 C5 98 8D B8 D9 E9 5C 6C 97 7E 83 42 88 " +
                            "     76 23 46 A1 F4 BE 19 61 EC 8A 52 28 E6 7A A6 11 " +
                            "     3A CC A4 B1 8A F9 20 ED 06 73 CB 76 21 5C D5 38 " +
                            "     8E A2 3B AC 58 A5 C3 85 92 35 E2 B1 14 C5 E7 F2 " +
                            "     5B 51 C8 B8 7C F9 18 54 B3 67 26 BF D2 F9 20 4F " +
                            "     C8 76 C1 09 A5 D7 57 6D FA 00 C8 8B D6 9F 97 1B " +
                            "     2B 17 71 96 AD DF 56 E9 69 A1 8B 9B 73 0A 03 40 " +
                            "     F2 69 BF B6 15 CF 99 F3 C2 52 08 FD 42 DB D0 14 " +
                            "     26 A1 42 5B A7 BA 28 8D 45 32 17 AE FA 09 FD 8E " +
                            "     BE B4 46 21 6B 76 14 86 80 0A 53 1F FC 83 7E 95 " +
                            "     D2 09 1F 5C 12 93 6E 70 32 E2 68 48 48 04 F8 10 " +
                            "     EA CD 11 ED 87 5B 1D 1B FF 61 13 72 FF 0F EE C7 " +
                            "     B7 49 B0 02 8C CF 4F EA C4 F9 8B 27 BF 00 00 00 " +
                            "     FF FF 03 00 50 4B 01 02 2D 00 14 00 06 00 08 00 " +
                            "     00 00 21 00 F0 F7 8A BB FD 00 00 00 E2 01 00 00 " +
                            "     13 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 00 5B 43 6F 6E 74 65 6E 74 5F 54 79 70 65 73 " +
                            "     5D 2E 78 6D 6C 50 4B 01 02 2D 00 14 00 06 00 08 " +
                            "     00 00 00 21 00 31 DD 5F 61 D2 00 00 00 8F 01 00 " +
                            "     00 0B 00 00 00 00 00 00 00 00 00 00 00 00 00 2E " +
                            "     01 00 00 5F 72 65 6C 73 2F 2E 72 65 6C 73 50 4B " +
                            "     01 02 2D 00 14 00 06 00 08 00 00 00 21 00 49 0D " +
                            "     41 2E 77 02 00 00 1A 07 00 00 10 00 00 00 00 00 " +
                            "     00 00 00 00 00 00 00 00 29 02 00 00 64 72 73 2F " +
                            "     73 68 61 70 65 78 6D 6C 2E 78 6D 6C 50 4B 01 02 " +
                            "     2D 00 14 00 06 00 08 00 00 00 21 00 79 45 76 AD " +
                            "     2B 01 00 00 AA 01 00 00 0F 00 00 00 00 00 00 00 " +
                            "     00 00 00 00 00 00 CE 04 00 00 64 72 73 2F 64 6F " +
                            "     77 6E 72 65 76 2E 78 6D 6C 50 4B 05 06 00 00 00 " +
                            "     00 04 00 04 00 F5 00 00 00 26 06 00 00 00 00 00 " +
                            "     00 10 F0 12 00 00 00 02 00 01 00 60 01 09 00 0D " +
                            "     00 02 00 D0 00 0A 00 73 00 00 00 11 F0 00 00 00 " +
                            "     00 5D 00 1A 00 15 00 12 00 06 00 11 00 11 60 00 " +
                            "     00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 EC " +
                            "     00 08 00 00 00 0D F0 00 00 00 00 B6 01 12 00 12 " +
                            "     02 00 00 00 00 00 00 00 00 06 00 10 00 00 00 00 " +
                            "     00 3C 00 07 00 00 74 65 78 74 2D 33 3C 00 10 00 " +
                            "     00 00 16 00 00 00 00 00 06 00 00 00 00 00 00 00 " +
                            "     EC 00 CC 07 0F 00 04 F0 CC 07 00 00 A2 0C 0A F0 " +
                            "     08 00 00 00 12 04 00 00 00 0A 00 00 93 00 0B F0 " +
                            "     4A 00 00 00 7F 00 00 00 04 00 80 00 00 36 9A 02 " +
                            "     85 00 02 00 00 00 8B 00 02 00 00 00 BF 00 1A 00 " +
                            "     1F 00 BF 01 01 00 11 00 FF 01 00 00 08 00 80 C3 " +
                            "     14 00 00 00 BF 03 00 00 02 00 54 00 65 00 78 00 " +
                            "     74 00 42 00 6F 00 78 00 20 00 34 00 00 00 13 00 " +
                            "     22 F1 38 07 00 00 A9 C3 32 07 00 00 50 4B 03 04 " +
                            "     14 00 06 00 08 00 00 00 21 00 F0 F7 8A BB FD 00 " +
                            "     00 00 E2 01 00 00 13 00 00 00 5B 43 6F 6E 74 65 " +
                            "     6E 74 5F 54 79 70 65 73 5D 2E 78 6D 6C 94 91 CD " +
                            "     4A C4 30 10 C7 EF 82 EF 10 E6 2A 6D AA 07 11 69 " +
                            "     BA 07 AB 47 15 5D 1F 60 48 A6 6D D8 36 09 99 58 " +
                            "     77 DF DE 74 3F 2E E2 0A 1E 67 E6 FF F1 23 A9 57 " +
                            "     DB 69 14 33 45 B6 DE 29 B8 2E 2B 10 E4 B4 37 D6 " +
                            "     F5 0A 3E D6 4F C5 1D 08 4E E8 0C 8E DE 91 82 1D " +
                            "     31 AC 9A CB 8B 7A BD 0B C4 22 BB 1D 2B 18 52 0A " +
                            "     F7 52 B2 1E 68 42 2E 7D 20 97 2F 9D 8F 13 A6 3C " +
                            "     C6 5E 06 D4 1B EC 49 DE 54 D5 AD D4 DE 25 72 A9 " +
                            "     48 4B 06 34 75 4B 1D 7E 8E 49 3C 6E F3 FA 40 12 " +
                            "     69 64 10 0F 07 E1 D2 A5 00 43 18 AD C6 94 49 E5 " +
                            "     EC CC 8F 96 E2 D8 50 66 E7 5E C3 83 0D 7C 95 31 " +
                            "     40 FE DA B0 5C CE 17 1C 7D 2F F9 69 A2 35 24 5E " +
                            "     31 A6 67 9C 32 86 34 91 25 0F 18 28 6B CA BF 53 " +
                            "     16 CC 89 0B DF 75 56 53 D9 46 7E 5F 7C 27 A8 73 " +
                            "     E1 C6 7F B9 48 F3 7F B3 DB 6C 7B A3 F9 94 2E F7 " +
                            "     3F D4 7C 03 00 00 FF FF 03 00 50 4B 03 04 14 00 " +
                            "     06 00 08 00 00 00 21 00 31 DD 5F 61 D2 00 00 00 " +
                            "     8F 01 00 00 0B 00 00 00 5F 72 65 6C 73 2F 2E 72 " +
                            "     65 6C 73 A4 90 C1 6A C3 30 0C 86 EF 83 BD 83 D1 " +
                            "     BD 71 DA 43 19 A3 4E 6F 85 5E 4B 07 BB 0A 5B 49 " +
                            "     4C 63 CB 58 26 6D DF BE A6 30 58 46 6F 3B EA 17 " +
                            "     FA 3E F1 EF F6 B7 30 A9 99 B2 78 8E 06 D6 4D 0B " +
                            "     8A A2 65 E7 E3 60 E0 EB 7C 58 7D 80 92 82 D1 E1 " +
                            "     C4 91 0C DC 49 60 DF BD BF ED 4E 34 61 A9 47 32 " +
                            "     FA 24 AA 52 A2 18 18 4B 49 9F 5A 8B 1D 29 A0 34 " +
                            "     9C 28 D6 4D CF 39 60 A9 63 1E 74 42 7B C1 81 F4 " +
                            "     A6 6D B7 3A FF 66 40 B7 60 AA A3 33 90 8F 6E 03 " +
                            "     EA 7C 4F D5 FC 87 1D BC CD 2C DC 97 C6 72 D0 DC " +
                            "     F7 DE BE A2 6A C7 D7 78 A2 B9 52 30 0F 54 0C B8 " +
                            "     2C CF 30 D3 DC D4 E7 40 BF F6 AE FF E9 95 11 13 " +
                            "     7D 57 FE 42 FC 4C AB F5 C7 AC 17 35 76 0F 00 00 " +
                            "     00 FF FF 03 00 50 4B 03 04 14 00 06 00 08 00 00 " +
                            "     00 21 00 5B 36 01 DE 77 02 00 00 1D 07 00 00 10 " +
                            "     00 00 00 64 72 73 2F 73 68 61 70 65 78 6D 6C 2E " +
                            "     78 6D 6C BC 55 4D 53 DB 30 10 BD 77 A6 FF 41 A3 " +
                            "     7B B0 9D 26 99 90 C1 66 28 0C BD 74 0A 93 C0 0F " +
                            "     10 B6 9C 78 2A 4B 1E 49 24 0E BF BE 6F 25 27 50 " +
                            "     0E 3D 90 0E 17 59 5E 69 F7 ED BE FD D0 C5 65 DF " +
                            "     2A B6 95 D6 35 46 E7 3C 3B 4B 39 93 BA 34 55 A3 " +
                            "     D7 39 7F 7C B8 1D CD 39 73 5E E8 4A 28 A3 65 CE " +
                            "     F7 D2 F1 CB E2 EB 97 8B BE B2 0B A1 CB 8D B1 0C " +
                            "     26 B4 5B 40 90 F3 8D F7 DD 22 49 5C B9 91 AD 70 " +
                            "     67 A6 93 1A A7 B5 B1 AD F0 F8 B5 EB A4 B2 62 07 " +
                            "     E3 AD 4A C6 69 3A 4B 5C 67 A5 A8 DC 46 4A 7F 13 " +
                            "     4F 78 11 6C 03 ED 5A 2A 75 15 20 A2 A8 B6 A6 8D " +
                            "     BB D2 A8 22 BB 48 C8 07 DA 06 05 6C EE EA BA 18 " +
                            "     A7 E7 D3 69 7A 3C 23 51 38 B6 66 57 64 83 0E ED " +
                            "     0F 42 BA 90 65 93 6F E9 A0 83 B3 A0 13 8C BF 22 " +
                            "     CA DE B3 B2 CF F9 74 3C 1D A7 E0 A4 DC E7 7C 3C " +
                            "     9B 4C 67 29 4F A2 2D D7 B1 56 94 D6 E4 9C 33 8F " +
                            "     EB AA D1 BF B1 8F 87 7A BB EA EE 6D DC 97 BF B6 " +
                            "     F7 96 35 15 8C 71 A6 45 0B 56 1F 70 FF BB E9 D9 " +
                            "     E4 60 0C 77 48 81 F9 1E 62 E4 85 E4 C1 A3 B7 86 " +
                            "     5C 30 29 16 7D 6D DB 21 09 E2 03 29 68 45 A3 E1 " +
                            "     A6 58 98 BA 66 00 9B 67 E7 D9 14 75 40 11 82 CD " +
                            "     79 1A 42 14 8B 7F 53 90 44 3F C8 4E 67 9D FF 21 " +
                            "     CD C9 3E 31 32 94 73 2B 4B 1F FC 13 DB 9F CE 13 " +
                            "     11 AF 10 04 A7 CD 6D A3 D4 A9 04 1C 08 8E A4 52 " +
                            "     65 39 BF 57 92 00 94 5E 4A 30 13 AA FC C3 04 23 " +
                            "     E1 E0 36 0D 81 B8 D2 AE 9F AE 95 65 E8 18 10 8D " +
                            "     4E C3 FA 44 6B 8C 2E 00 12 72 8D C0 3E 19 7B 80 " +
                            "     24 74 59 D7 A0 FE 93 F1 8F A0 21 7E A3 FF 1F 7E " +
                            "     DB 68 63 07 FE 31 9E 24 25 60 2B 54 CE 7D 1F FB " +
                            "     0B 7C 47 BC A1 D7 86 02 A0 5A A0 3E AC F6 E4 D2 " +
                            "     13 BE 68 CC 53 AB 01 23 D7 DF 61 A9 95 D9 E5 BC " +
                            "     54 4D C7 19 66 E9 CB 7B D9 CE 8A 2E E7 1A D3 90 " +
                            "     33 EB D5 B5 81 BF 28 96 38 79 E1 7A 8C A7 BB 7A " +
                            "     F6 E8 82 A1 39 A2 8B E4 AC 72 7E 45 55 7C AA BB " +
                            "     28 4B F4 DC A9 56 C8 88 0D 0B 08 54 82 5E 18 A9 " +
                            "     47 8F 2B BC 30 2F 98 71 D9 61 D0 F8 82 E6 E7 68 " +
                            "     42 8D EE 43 BB 07 2D A9 AB 7B 61 C5 F2 A8 6C 9F " +
                            "     47 CB C7 BF 95 69 34 0C D9 3B A4 2C 8C 4D 07 69 " +
                            "     78 2D 54 23 B5 BF 11 5E 1C FA FD DD 3B 13 6E 47 " +
                            "     76 8B 3F 00 00 00 FF FF 03 00 50 4B 03 04 14 00 " +
                            "     06 00 08 00 00 00 21 00 DB CE 2D F9 2C 01 00 00 " +
                            "     AE 01 00 00 0F 00 00 00 64 72 73 2F 64 6F 77 6E " +
                            "     72 65 76 2E 78 6D 6C 5C 50 5D 4F C2 30 14 7D 37 " +
                            "     F1 3F 34 D7 C4 37 E9 18 D4 00 52 C8 62 42 C4 64 " +
                            "     7E 0C 49 7C 2D 6B CB 16 D7 76 B6 75 0C 7F BD 1D " +
                            "     42 30 3E DD 9C 73 EE 39 ED B9 D3 79 AB 2A D4 08 " +
                            "     EB 4A A3 29 F4 7B 11 20 A1 73 C3 4B BD A5 B0 7E " +
                            "     5B DC 8C 00 39 CF 34 67 95 D1 82 C2 5E 38 98 CF " +
                            "     2E 2F A6 6C C2 CD 4E 67 A2 59 F9 2D 0A 21 DA 4D " +
                            "     18 85 C2 FB 7A 82 B1 CB 0B A1 98 EB 99 5A E8 A0 " +
                            "     49 63 15 F3 01 DA 2D E6 96 ED 42 B8 AA 70 1C 45 " +
                            "     B7 58 B1 52 87 17 0A 56 8B FB 42 E4 1F AB 2F 45 " +
                            "     21 E5 F2 E9 7D 58 2E 1E 99 FC 4C 96 C3 6C 44 9A " +
                            "     17 4E 28 BD BE 6A 93 3B 40 5E B4 FE BC 5C 37 36 " +
                            "     8D D3 74 F3 BA 4E 4E 0B 87 B8 25 A7 40 00 C9 87 " +
                            "     FD C6 96 3C 63 CE 0B 4B 21 F4 0B 6D 43 53 98 85 " +
                            "     0A 6D 95 E8 BC 30 16 C9 4C B8 F2 3B F4 FB E5 A5 " +
                            "     35 0A 59 B3 0B B8 0F 28 37 55 27 74 C4 B3 94 4E " +
                            "     F8 8E 1E 0E A2 90 15 A4 13 15 47 63 42 22 C0 5D " +
                            "     AE 37 47 77 7C 74 87 F9 D7 3D 8E C8 3F 77 3F 1E " +
                            "     8C 62 D2 B9 F1 F9 5B 07 70 3E F3 EC 07 00 00 FF " +
                            "     FF 03 00 50 4B 01 02 2D 00 14 00 06 00 08 00 00 " +
                            "     00 21 00 F0 F7 8A BB FD 00 00 00 E2 01 00 00 13 " +
                            "     00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 5B 43 6F 6E 74 65 6E 74 5F 54 79 70 65 73 5D " +
                            "     2E 78 6D 6C 50 4B 01 02 2D 00 14 00 06 00 08 00 " +
                            "     00 00 21 00 31 DD 5F 61 D2 00 00 00 8F 01 00 00 " +
                            "     0B 00 00 00 00 00 00 00 00 00 00 00 00 00 2E 01 " +
                            "     00 00 5F 72 65 6C 73 2F 2E 72 65 6C 73 50 4B 01 " +
                            "     02 2D 00 14 00 06 00 08 00 00 00 21 00 5B 36 01 " +
                            "     DE 77 02 00 00 1D 07 00 00 10 00 00 00 00 00 00 " +
                            "     00 00 00 00 00 00 00 29 02 00 00 64 72 73 2F 73 " +
                            "     68 61 70 65 78 6D 6C 2E 78 6D 6C 50 4B 01 02 2D " +
                            "     00 14 00 06 00 08 00 00 00 21 00 DB CE 2D F9 2C " +
                            "     01 00 00 AE 01 00 00 0F 00 00 00 00 00 00 00 00 " +
                            "     00 00 00 00 00 CE 04 00 00 64 72 73 2F 64 6F 77 " +
                            "     6E 72 65 76 2E 78 6D 6C 50 4B 05 06 00 00 00 00 " +
                            "     04 00 04 00 F5 00 00 00 27 06 00 00 00 00 00 00 " +
                            "     10 F0 12 00 00 00 02 00 01 00 60 01 0B 00 9A 00 " +
                            "     02 00 D0 00 0C 00 00 01 00 00 11 F0 00 00 00 00 " +
                            "     5D 00 1A 00 15 00 12 00 06 00 12 00 11 60 00 00 " +
                            "     00 00 00 00 00 00 00 00 00 00 00 00 00 00 EC 00 " +
                            "     08 00 00 00 0D F0 00 00 00 00 B6 01 12 00 12 02 " +
                            "     00 00 00 00 00 00 00 00 06 00 10 00 00 00 00 00 " +
                            "     3C 00 07 00 00 74 65 78 74 2D 34 3C 00 10 00 00 " +
                            "     00 16 00 00 00 00 00 06 00 00 00 00 00 00 00 3C " +
                            "     00 C8 07 0F 00 04 F0 C8 07 00 00 A2 0C 0A F0 08 " +
                            "     00 00 00 13 04 00 00 00 0A 00 00 93 00 0B F0 4A " +
                            "     00 00 00 7F 00 00 00 04 00 80 00 80 33 9A 02 85 " +
                            "     00 02 00 00 00 8B 00 02 00 00 00 BF 00 1A 00 1F " +
                            "     00 BF 01 01 00 11 00 FF 01 00 00 08 00 80 C3 14 " +
                            "     00 00 00 BF 03 00 00 02 00 54 00 65 00 78 00 74 " +
                            "     00 42 00 6F 00 78 00 20 00 35 00 00 00 13 00 22 " +
                            "     F1 34 07 00 00 A9 C3 2E 07 00 00 50 4B 03 04 14 " +
                            "     00 06 00 08 00 00 00 21 00 F0 F7 8A BB FD 00 00 " +
                            "     00 E2 01 00 00 13 00 00 00 5B 43 6F 6E 74 65 6E " +
                            "     74 5F 54 79 70 65 73 5D 2E 78 6D 6C 94 91 CD 4A " +
                            "     C4 30 10 C7 EF 82 EF 10 E6 2A 6D AA 07 11 69 BA " +
                            "     07 AB 47 15 5D 1F 60 48 A6 6D D8 36 09 99 58 77 " +
                            "     DF DE 74 3F 2E E2 0A 1E 67 E6 FF F1 23 A9 57 DB " +
                            "     69 14 33 45 B6 DE 29 B8 2E 2B 10 E4 B4 37 D6 F5 " +
                            "     0A 3E D6 4F C5 1D 08 4E E8 0C 8E DE 91 82 1D 31 " +
                            "     AC 9A CB 8B 7A BD 0B C4 22 BB 1D 2B 18 52 0A F7 " +
                            "     52 B2 1E 68 42 2E 7D 20 97 2F 9D 8F 13 A6 3C C6 " +
                            "     5E 06 D4 1B EC 49 DE 54 D5 AD D4 DE 25 72 A9 48 " +
                            "     4B 06 34 75 4B 1D 7E 8E 49 3C 6E F3 FA 40 12 69 " +
                            "     64 10 0F 07 E1 D2 A5 00 43 18 AD C6 94 49 E5 EC " +
                            "     CC 8F 96 E2 D8 50 66 E7 5E C3 83 0D 7C 95 31 40 " +
                            "     FE DA B0 5C CE 17 1C 7D 2F F9 69 A2 35 24 5E 31 " +
                            "     A6 67 9C 32 86 34 91 25 0F 18 28 6B CA BF 53 16 " +
                            "     CC 89 0B DF 75 56 53 D9 46 7E 5F 7C 27 A8 73 E1 " +
                            "     C6 7F B9 48 F3 7F B3 DB 6C 7B A3 F9 94 2E F7 3F " +
                            "     D4 7C 03 00 00 FF FF 03 00 50 4B 03 04 14 00 06 " +
                            "     00 08 00 00 00 21 00 31 DD 5F 61 D2 00 00 00 8F " +
                            "     01 00 00 0B 00 00 00 5F 72 65 6C 73 2F 2E 72 65 " +
                            "     6C 73 A4 90 C1 6A C3 30 0C 86 EF 83 BD 83 D1 BD " +
                            "     71 DA 43 19 A3 4E 6F 85 5E 4B 07 BB 0A 5B 49 4C " +
                            "     63 CB 58 26 6D DF BE A6 30 58 46 6F 3B EA 17 FA " +
                            "     3E F1 EF F6 B7 30 A9 99 B2 78 8E 06 D6 4D 0B 8A " +
                            "     A2 65 E7 E3 60 E0 EB 7C 58 7D 80 92 82 D1 E1 C4 " +
                            "     91 0C DC 49 60 DF BD BF ED 4E 34 61 A9 47 32 FA " +
                            "     24 AA 52 A2 18 18 4B 49 9F 5A 8B 1D 29 A0 34 9C " +
                            "     28 D6 4D CF 39 60 A9 63 1E 74 42 7B C1 81 F4 A6 " +
                            "     6D B7 3A FF 66 40 B7 60 AA A3 33 90 8F 6E 03 EA " +
                            "     7C 4F D5 FC 87 1D BC CD 2C DC 97 C6 72 D0 DC F7 " +
                            "     DE BE A2 6A C7 D7 78 A2 B9 52 30 0F 54 0C B8 2C " +
                            "     CF 30 D3 DC D4 E7 40 BF F6 AE FF E9 95 11 13 7D " +
                            "     57 FE 42 FC 4C AB F5 C7 AC 17 35 76 0F 00 00 00 " +
                            "     FF FF 03 00 50 4B 03 04 14 00 06 00 08 00 00 00 " +
                            "     21 00 C4 6F 3F 83 76 02 00 00 1C 07 00 00 10 00 " +
                            "     00 00 64 72 73 2F 73 68 61 70 65 78 6D 6C 2E 78 " +
                            "     6D 6C BC 55 CD 6E DB 30 0C BE 0F D8 3B 08 BA A7 " +
                            "     B6 83 3A 4D 8D C8 45 97 A2 BB 0C 6B 90 34 0F A0 " +
                            "     C8 72 62 4C 96 0C 49 4D 9C 3E FD 48 C9 49 B7 1E " +
                            "     76 68 86 5E 64 9A 12 F9 51 1F 7F 34 BB EB 5B 45 " +
                            "     F6 D2 BA C6 68 46 B3 AB 94 12 A9 85 A9 1A BD 65 " +
                            "     74 FD FC 38 9A 52 E2 3C D7 15 57 46 4B 46 8F D2 " +
                            "     D1 BB F2 EB 97 59 5F D9 82 6B B1 33 96 80 0B ED " +
                            "     0A 50 30 BA F3 BE 2B 92 C4 89 9D 6C B9 BB 32 9D " +
                            "     D4 B0 5B 1B DB 72 0F BF 76 9B 54 96 1F C0 79 AB " +
                            "     92 71 9A 4E 12 D7 59 C9 2B B7 93 D2 3F C4 1D 5A " +
                            "     06 DF 80 36 97 4A DD 07 88 A8 AA AD 69 A3 24 8C " +
                            "     2A B3 59 82 31 A0 18 0C 40 78 AA EB 72 9C DE E6 " +
                            "     79 7A DE 43 55 D8 B6 E6 50 66 D7 51 8F F2 49 19 " +
                            "     6C A6 F9 4D 7E DE 0A 26 C1 F7 1B A0 EC 3D 11 3D " +
                            "     A3 F9 38 1F A7 40 89 38 32 3A 9E 5C E7 93 94 26 " +
                            "     D1 95 EB 48 CB 85 35 8C 52 E2 E1 B8 6A F4 2F 90 " +
                            "     E3 A6 DE AF BA 85 8D B2 F8 B9 5F 58 D2 54 8C 4E " +
                            "     28 D1 BC 05 52 9F E1 FC 37 D3 93 FC E4 0C CE A0 " +
                            "     01 F1 3D A8 21 2D A8 0F 11 FD E9 C8 05 97 BC E8 " +
                            "     6B DB 0E 39 E0 1F C8 40 CB 1B 0D 61 F2 C2 D4 35 " +
                            "     01 B0 69 76 9B E5 50 06 E1 86 C0 E5 4D 88 8A 17 " +
                            "     FF A6 20 89 71 A0 9F CE 3A FF 5D 9A 8B 63 22 E8 " +
                            "     88 51 2B 85 0F F1 F1 FD 0F E7 91 88 37 08 84 D3 " +
                            "     E6 B1 51 EA 52 02 4E 04 47 52 B1 B0 9C 3F 2A 89 " +
                            "     00 4A 2F 25 30 13 8A FC C3 04 43 C2 81 DB 34 5C " +
                            "     C4 09 BB DD CC 95 25 D0 30 40 34 34 1A AC 1B 5C " +
                            "     E3 ED 02 20 22 D7 70 B1 4F C6 1E 20 11 5D D6 35 " +
                            "     50 FF C9 F8 67 D0 70 7F A3 FF 1F 7E DB 68 63 07 " +
                            "     FE 61 3A 49 4C C0 9E 2B 46 7D 1F FB 0B F8 8E 78 " +
                            "     43 AF 0D 05 80 B5 80 7D 58 1D 31 A4 0D 7C A1 31 " +
                            "     2F AD 06 98 B8 FE 09 96 5A 99 03 A3 42 35 1D 25 " +
                            "     30 4A 5F DF EB 0E 96 77 8C 6A 18 86 94 58 AF E6 " +
                            "     06 E2 85 62 89 83 17 42 8F F7 E9 EE 5F 3C 74 C1 " +
                            "     D0 1C 31 44 0C 56 39 BF C2 2A BE 34 5C 28 4B E8 " +
                            "     B9 4B BD A0 13 1B 16 20 50 71 7C 60 A4 1E AD 57 " +
                            "     F0 C0 BC C2 8C CB D2 50 FF BC F0 25 CE CF 11 0C " +
                            "     64 94 71 0D 56 52 57 0B 6E F9 F2 6C 6C 5F 46 CB " +
                            "     F5 DF C6 38 1A 86 EC 9D 52 16 C6 A6 03 6D 78 2C " +
                            "     54 23 B5 7F E0 9E 9F FA FD DD 33 13 4E 47 76 CB " +
                            "     DF 00 00 00 FF FF 03 00 50 4B 03 04 14 00 06 00 " +
                            "     08 00 00 00 21 00 AF 34 70 A0 29 01 00 00 AD 01 " +
                            "     00 00 0F 00 00 00 64 72 73 2F 64 6F 77 6E 72 65 " +
                            "     76 2E 78 6D 6C 5C 50 5D 4F C2 30 14 7D 37 F1 3F " +
                            "     34 D7 C4 37 D9 98 2B 4C A4 90 C5 84 88 C9 FC 18 " +
                            "     92 F0 5A D6 96 2D AE ED 6C 2B 0C 7F BD 1D A8 44 " +
                            "     1F CF 67 EF E9 78 DA CA 1A 6D B9 B1 95 56 04 FA " +
                            "     BD 10 10 57 85 66 95 DA 10 58 BE CE AE 12 40 D6 " +
                            "     51 C5 68 AD 15 27 B0 E7 16 A6 93 F3 B3 31 1D 31 " +
                            "     BD 53 39 DF 2E DC 06 F9 12 65 47 94 40 E9 5C 33 " +
                            "     0A 02 5B 94 5C 52 DB D3 0D 57 5E 13 DA 48 EA 3C " +
                            "     34 9B 80 19 BA F3 E5 B2 0E A2 30 1C 04 92 56 CA " +
                            "     BF 50 D2 86 DF 95 BC 78 5B 7C 48 02 19 13 8F AB " +
                            "     B8 9A 3D 50 F1 9E CE E3 3C C1 DB 67 86 09 B9 BC " +
                            "     68 D3 5B 40 8E B7 EE 64 6E EA 55 16 65 D9 FA 65 " +
                            "     99 FE 18 0E 75 73 46 60 00 48 DC EF D7 A6 62 39 " +
                            "     B5 8E 1B 02 7E 9F 5F EB 97 C2 C4 4F 68 EB 54 15 " +
                            "     A5 36 48 E4 DC 56 9F 7E DF 91 17 46 4B 64 F4 CE " +
                            "     E3 18 50 A1 EB 4E E8 88 27 21 2C 77 04 A2 04 0F " +
                            "     F1 41 F9 65 C2 1B 8C 43 08 BA 5A A7 BF C3 47 8B " +
                            "     B7 FF 09 F7 C3 78 F8 2F DD 8F AE 93 08 77 E9 E0 " +
                            "     74 D5 01 9C 7E 79 F2 05 00 00 FF FF 03 00 50 4B " +
                            "     01 02 2D 00 14 00 06 00 08 00 00 00 21 00 F0 F7 " +
                            "     8A BB FD 00 00 00 E2 01 00 00 13 00 00 00 00 00 " +
                            "     00 00 00 00 00 00 00 00 00 00 00 00 5B 43 6F 6E " +
                            "     74 65 6E 74 5F 54 79 70 65 73 5D 2E 78 6D 6C 50 " +
                            "     4B 01 02 2D 00 14 00 06 00 08 00 00 00 21 00 31 " +
                            "     DD 5F 61 D2 00 00 00 8F 01 00 00 0B 00 00 00 00 " +
                            "     00 00 00 00 00 00 00 00 00 2E 01 00 00 5F 72 65 " +
                            "     6C 73 2F 2E 72 65 6C 73 50 4B 01 02 2D 00 14 00 " +
                            "     06 00 08 00 00 00 21 00 C4 6F 3F 83 76 02 00 00 " +
                            "     1C 07 00 00 10 00 00 00 00 00 00 00 00 00 00 00 " +
                            "     00 00 29 02 00 00 64 72 73 2F 73 68 61 70 65 78 " +
                            "     6D 6C 2E 78 6D 6C 50 4B 01 02 2D 00 14 00 06 00 " +
                            "     08 00 00 00 21 00 AF 34 70 A0 29 01 00 00 AD 01 " +
                            "     00 00 0F 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                            "     CD 04 00 00 64 72 73 2F 64 6F 77 6E 72 65 76 2E " +
                            "     78 6D 6C 50 4B 05 06 00 00 00 00 04 00 04 00 F5 " +
                            "     00 00 00 23 06 00 00 00 00 00 00 10 F0 12 00 00 " +
                            "     00 02 00 01 00 60 01 0E 00 26 00 02 00 D0 00 0F " +
                            "     00 8D 00 00 00 11 F0 00 00 00 00 5D 00 1A 00 15 " +
                            "     00 12 00 06 00 13 00 11 60 00 00 00 00 00 00 00 " +
                            "     00 00 00 00 00 00 00 00 00 3C 00 08 00 00 00 0D " +
                            "     F0 00 00 00 00 B6 01 12 00 12 02 00 00 00 00 00 " +
                            "     00 00 00 06 00 10 00 00 00 00 00 3C 00 07 00 00 " +
                            "     74 65 78 74 2D 35 3C 00 10 00 00 00 16 00 00 00 " +
                            "     00 00 06 00 00 00 00 00 00 00  " +
                            "     ";

            byte[] dgBytes = HexRead.ReadFromString(data);
            IList<NPOI.HSSF.Record.Record> dgRecords = RecordFactory.CreateRecords(new MemoryStream(dgBytes));
            Assert.AreEqual(20, dgRecords.Count);

            short[] expectedSids = {
                    DrawingRecord.sid,
                    ObjRecord.sid,
                    DrawingRecord.sid,
                    TextObjectRecord.sid,
                    DrawingRecord.sid,
                    ObjRecord.sid,
                    DrawingRecord.sid,
                    TextObjectRecord.sid,
                    DrawingRecord.sid,
                    ObjRecord.sid,
                    DrawingRecord.sid,
                    TextObjectRecord.sid,
                    DrawingRecord.sid,
                    ObjRecord.sid,
                    DrawingRecord.sid,
                    TextObjectRecord.sid,
                    ContinueRecord.sid,
                    ObjRecord.sid,
                    ContinueRecord.sid,
                    TextObjectRecord.sid
            };
            for (int i = 0; i < expectedSids.Length; i++)
            {
                Assert.AreEqual(expectedSids[i], dgRecords[(i)].Sid, "unexpected record.sid and index[" + i + "]");
            }
            DrawingManager2 drawingManager = new DrawingManager2(new EscherDggRecord());

            // create a dummy sheet consisting of our Test data
            InternalSheet sheet = InternalSheet.CreateSheet();
            List<RecordBase> records = sheet.Records;
            records.Clear();
            records.AddRange(dgRecords);
            records.Add(EOFRecord.instance);


            sheet.AggregateDrawingRecords(drawingManager, false);
            Assert.AreEqual(2, records.Count, "drawing was not fully aggregated");
            Assert.IsTrue(records[(0)] is EscherAggregate, "expected EscherAggregate");
            Assert.IsTrue(records[(1)] is EOFRecord, "expected EOFRecord");
            EscherAggregate agg = (EscherAggregate)records[(0)];

            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length, "different size of Drawing data before and After save");
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data before and After save is different");
        }
        [Test]
        public void TestUnhandledContinue2()
        {
            String data = "EC 00 38 08 0F 00 02 F0 97 37 00 00 10 00 " +
                    "08 F0 08 00 00 00 08 00 00 00 07 04 00 00 0F 00 " +
                    "03 F0 7F 37 00 00 0F 00 04 F0 28 00 00 00 01 00 " +
                    "09 F0 10 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 00 02 00 0A F0 08 00 00 00 00 04 " +
                    "00 00 05 00 00 00 0F 00 04 F0 E0 07 00 00 12 00 " +
                    "0A F0 08 00 00 00 01 04 00 00 00 0A 00 00 83 00 " +
                    "0B F0 50 00 00 00 BF 00 18 00 1F 00 81 01 4F 81 " +
                    "BD 00 BF 01 10 00 10 00 C0 01 38 5D 8A 00 CB 01 " +
                    "38 63 00 00 FF 01 08 00 08 00 80 C3 20 00 00 00 " +
                    "BF 03 00 00 02 00 1F 04 40 04 4F 04 3C 04 3E 04 " +
                    "43 04 33 04 3E 04 3B 04 4C 04 3D 04 38 04 3A 04 " +
                    "20 00 31 00 00 00 23 00 22 F1 4E 07 00 00 FF 01 " +
                    "00 00 40 00 A9 C3 42 07 00 00 50 4B 03 04 14 00 " +
                    "06 00 08 00 00 00 21 00 F0 F7 8A BB FD 00 00 00 " +
                    "E2 01 00 00 13 00 00 00 5B 43 6F 6E 74 65 6E 74 " +
                    "5F 54 79 70 65 73 5D 2E 78 6D 6C 94 91 CD 4A C4 " +
                    "30 10 C7 EF 82 EF 10 E6 2A 6D AA 07 11 69 BA 07 " +
                    "AB 47 15 5D 1F 60 48 A6 6D D8 36 09 99 58 77 DF " +
                    "DE 74 3F 2E E2 0A 1E 67 E6 FF F1 23 A9 57 DB 69 " +
                    "14 33 45 B6 DE 29 B8 2E 2B 10 E4 B4 37 D6 F5 0A " +
                    "3E D6 4F C5 1D 08 4E E8 0C 8E DE 91 82 1D 31 AC " +
                    "9A CB 8B 7A BD 0B C4 22 BB 1D 2B 18 52 0A F7 52 " +
                    "B2 1E 68 42 2E 7D 20 97 2F 9D 8F 13 A6 3C C6 5E " +
                    "06 D4 1B EC 49 DE 54 D5 AD D4 DE 25 72 A9 48 4B " +
                    "06 34 75 4B 1D 7E 8E 49 3C 6E F3 FA 40 12 69 64 " +
                    "10 0F 07 E1 D2 A5 00 43 18 AD C6 94 49 E5 EC CC " +
                    "8F 96 E2 D8 50 66 E7 5E C3 83 0D 7C 95 31 40 FE " +
                    "DA B0 5C CE 17 1C 7D 2F F9 69 A2 35 24 5E 31 A6 " +
                    "67 9C 32 86 34 91 25 0F 18 28 6B CA BF 53 16 CC " +
                    "89 0B DF 75 56 53 D9 46 7E 5F 7C 27 A8 73 E1 C6 " +
                    "7F B9 48 F3 7F B3 DB 6C 7B A3 F9 94 2E F7 3F D4 " +
                    "7C 03 00 00 FF FF 03 00 50 4B 03 04 14 00 06 00 " +
                    "08 00 00 00 21 00 31 DD 5F 61 D2 00 00 00 8F 01 " +
                    "00 00 0B 00 00 00 5F 72 65 6C 73 2F 2E 72 65 6C " +
                    "73 A4 90 C1 6A C3 30 0C 86 EF 83 BD 83 D1 BD 71 " +
                    "DA 43 19 A3 4E 6F 85 5E 4B 07 BB 0A 5B 49 4C 63 " +
                    "CB 58 26 6D DF BE A6 30 58 46 6F 3B EA 17 FA 3E " +
                    "F1 EF F6 B7 30 A9 99 B2 78 8E 06 D6 4D 0B 8A A2 " +
                    "65 E7 E3 60 E0 EB 7C 58 7D 80 92 82 D1 E1 C4 91 " +
                    "0C DC 49 60 DF BD BF ED 4E 34 61 A9 47 32 FA 24 " +
                    "AA 52 A2 18 18 4B 49 9F 5A 8B 1D 29 A0 34 9C 28 " +
                    "D6 4D CF 39 60 A9 63 1E 74 42 7B C1 81 F4 A6 6D " +
                    "B7 3A FF 66 40 B7 60 AA A3 33 90 8F 6E 03 EA 7C " +
                    "4F D5 FC 87 1D BC CD 2C DC 97 C6 72 D0 DC F7 DE " +
                    "BE A2 6A C7 D7 78 A2 B9 52 30 0F 54 0C B8 2C CF " +
                    "30 D3 DC D4 E7 40 BF F6 AE FF E9 95 11 13 7D 57 " +
                    "FE 42 FC 4C AB F5 C7 AC 17 35 76 0F 00 00 00 FF " +
                    "FF 03 00 50 4B 03 04 14 00 06 00 08 00 00 00 21 " +
                    "00 B5 19 FD 5B 97 02 00 00 FC 06 00 00 10 00 00 " +
                    "00 64 72 73 2F 73 68 61 70 65 78 6D 6C 2E 78 6D " +
                    "6C AC 55 49 6E DB 30 14 DD 17 E8 1D 08 EE 13 0D " +
                    "B6 12 47 B0 14 B4 0E DA 4D D1 18 4E 73 00 56 A2 " +
                    "6C A1 14 29 90 AC 2D 67 55 A0 DB 02 3D 42 0F D1 " +
                    "4D D1 21 67 90 6F D4 4F 52 52 DA 74 58 C4 F6 42 " +
                    "A6 DE 17 FF 7B 7F 22 A7 E7 4D C5 D0 9A 4A 55 0A " +
                    "9E E0 E0 D8 C7 88 F2 4C E4 25 5F 26 F8 FA D5 B3 " +
                    "A3 09 46 4A 13 9E 13 26 38 4D F0 96 2A 7C 9E 3E " +
                    "7E 34 6D 72 19 13 9E AD 84 44 E0 82 AB 18 80 04 " +
                    "AF B4 AE 63 CF 53 D9 8A 56 44 1D 8B 9A 72 B0 16 " +
                    "42 56 44 C3 AB 5C 7A B9 24 1B 70 5E 31 2F F4 FD " +
                    "13 4F D5 92 92 5C AD 28 D5 17 CE 82 53 EB 5B 6F " +
                    "C4 8C 32 F6 C4 52 38 A8 90 A2 72 AB 4C B0 34 9C " +
                    "7A 46 83 59 DA 0D B0 B8 2C 8A 74 14 8C 47 61 34 " +
                    "D8 0C 64 CD 52 6C D2 91 83 CD B2 C7 8C 3D 88 C2 " +
                    "B1 EF 0F 36 BB C5 FA BE 23 D4 62 20 49 C7 83 F3 " +
                    "01 33 5B A2 49 E0 FF 8B 38 E8 C4 DE 67 3E 8B FA " +
                    "1D 60 B9 E3 ED D9 54 8D 2A 92 49 91 60 8C 34 6D " +
                    "34 2B F9 1B 58 3B 5A BE BE AA E7 B2 93 F0 72 3D " +
                    "97 A8 CC 13 1C 62 C4 49 05 85 6A 3F ED DE ED 3E " +
                    "B6 DF DB DB DD FB F6 73 7B DB 7E DB 7D 68 7F B4 " +
                    "5F DA AF 28 C0 DE B0 CD F8 80 37 1B EE AF 1E 95 " +
                    "F5 4D E2 A6 90 55 57 60 F2 80 F2 56 A4 E4 A0 97 " +
                    "C4 A2 28 50 03 0D 16 8D 46 10 33 46 DB 04 9F 86 " +
                    "A3 33 DF 37 5A 48 0C C1 A1 CC D8 C7 93 C8 80 28 " +
                    "83 0F 82 E8 34 38 81 8F 8D 3E A7 C4 7C 5A 4B A5 " +
                    "9F 53 B1 B7 2A 64 1C 25 58 D2 4C 5B 85 64 FD 42 " +
                    "69 47 D5 53 74 79 71 B9 30 CD A6 F4 96 51 23 82 " +
                    "F1 05 85 80 6C E3 3F 38 2F 50 30 08 39 B4 EC 76 " +
                    "62 E8 8C 49 B4 26 2C C1 24 CB 28 D7 81 33 AD 48 " +
                    "4E 1D 1C F9 F0 EB F2 31 EC B0 D9 B1 82 8C B2 A2 " +
                    "64 EC 60 DA 3A 01 66 9A FF D4 E6 72 D5 F1 D9 22 " +
                    "16 05 24 F3 60 E4 FE FF 12 E3 C8 69 CF 68 23 17 " +
                    "FC 70 E4 55 C9 85 FC 9B 00 06 55 E9 22 77 7C 7D " +
                    "93 B8 D6 30 5D A2 9B A7 22 DF 1A 49 AF E1 1F 06 " +
                    "73 DF 3E 81 F3 59 5F C2 A3 60 62 93 E0 8C 95 35 " +
                    "46 70 F0 DE DC C7 A4 66 33 01 DD 03 F3 E3 8E E6 " +
                    "04 6B 37 5F 4C E9 2B 23 70 5F 29 10 39 4C E0 BE " +
                    "5E AC 13 C8 0B 61 4B B8 74 98 93 48 79 3E 27 92 " +
                    "2C 00 67 C4 DC 3E F2 ED D1 E2 1A 6E 9F 1B 38 09 " +
                    "82 A1 ED EB 2E DF 7D 92 ED C9 A5 00 B5 97 01 2B " +
                    "61 6C 2E 88 26 A6 44 B6 16 BF 5F 23 16 73 B9 49 " +
                    "7F 02 00 00 FF FF 03 00 50 4B 03 04 14 00 06 00 " +
                    "08 00 00 00 21 00 CA 39 EE E5 1C 01 00 00 8E 01 " +
                    "00 00 0F 00 00 00 64 72 73 2F 64 6F 77 6E 72 65 " +
                    "76 2E 78 6D 6C 4C 90 CB 4E C3 30 10 45 F7 48 FC " +
                    "83 35 48 6C 10 75 9E 28 94 3A 55 41 42 65 53 44 " +
                    "DA B2 60 67 12 E7 21 62 3B B2 4D 93 FE 3D 93 96 " +
                    "AA DD F9 CE DC 33 33 D7 B3 F9 20 5B B2 13 C6 36 " +
                    "5A 31 F0 27 1E 10 A1 72 5D 34 AA 62 B0 DD BC DE " +
                    "27 40 AC E3 AA E0 AD 56 82 C1 5E 58 98 A7 D7 57 " +
                    "33 3E 2D 74 AF 32 B1 5B BB 8A E0 10 65 A7 9C 41 " +
                    "ED 5C 37 A5 D4 E6 B5 90 DC 4E 74 27 14 F6 4A 6D " +
                    "24 77 28 4D 45 0B C3 7B 1C 2E 5B 1A 78 DE 03 95 " +
                    "BC 51 B8 A1 E6 9D 78 A9 45 FE B3 FE 95 B8 E4 43 " +
                    "7E 6E F5 73 F2 B5 A2 77 DB 3E 5B 6E 92 26 0E 13 " +
                    "C6 6E 6F 86 C5 13 10 27 06 77 36 FF D3 6F 05 83 " +
                    "00 48 B9 DC 7F 9B A6 C8 B8 75 C2 30 C0 38 18 0E " +
                    "83 41 8A 17 0F ED 42 E5 B5 36 E3 BB 34 5A 12 A3 " +
                    "7B 06 21 90 5C B7 07 1A F5 7B 59 5A E1 90 88 83 " +
                    "C8 43 1C 5B A7 52 E8 47 61 10 03 1D 71 A7 8F B0 " +
                    "8F 3B 0F 74 04 63 E1 64 7D 8C D1 78 C9 C6 89 EF " +
                    "1D 59 7A 79 07 8A F3 37 A6 7F 00 00 00 FF FF 03 " +
                    "00 50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 21 " +
                    "00 F0 F7 8A BB FD 00 00 00 E2 01 00 00 13 00 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 5B " +
                    "43 6F 6E 74 65 6E 74 5F 54 79 70 65 73 5D 2E 78 " +
                    "6D 6C 50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 " +
                    "21 00 31 DD 5F 61 D2 00 00 00 8F 01 00 00 0B 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 2E 01 00 00 " +
                    "5F 72 65 6C 73 2F 2E 72 65 6C 73 50 4B 01 02 2D " +
                    "00 14 00 06 00 08 00 00 00 21 00 B5 19 FD 5B 97 " +
                    "02 00 00 FC 06 00 00 10 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 29 02 00 00 64 72 73 2F 73 68 61 " +
                    "70 65 78 6D 6C 2E 78 6D 6C 50 4B 01 02 2D 00 14 " +
                    "00 06 00 08 00 00 00 21 00 CA 39 EE E5 1C 01 00 " +
                    "00 8E 01 00 00 0F 00 00 00 00 00 00 00 00 00 00 " +
                    "00 00 00 EE 04 00 00 64 72 73 2F 64 6F 77 6E 72 " +
                    "65 76 2E 78 6D 6C 50 4B 05 06 00 00 00 00 04 00 " +
                    "04 00 F5 00 00 00 37 06 00 00 00 00 00 00 10 F0 " +
                    "12 00 00 00 00 00 02 00 10 02 03 00 CD 00 04 00 " +
                    "D0 03 0C 00 0D 00 00 00 11 F0 00 00 00 00 5D 00 " +
                    "1A 00 15 00 12 00 02 00 01 00 11 60 00 00 00 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 EC 00 E8 07 " +
                    "0F 00 04 F0 E0 07 00 00 12 00 0A F0 08 00 00 00 " +
                    "02 04 00 00 00 0A 00 00 83 00 0B F0 50 00 00 00 " +
                    "BF 00 18 00 1F 00 81 01 4F 81 BD 00 BF 01 10 00 " +
                    "10 00 C0 01 38 5D 8A 00 CB 01 38 63 00 00 FF 01 " +
                    "08 00 08 00 80 C3 20 00 00 00 BF 03 00 00 02 00 " +
                    "1F 04 40 04 4F 04 3C 04 3E 04 43 04 33 04 3E 04 " +
                    "3B 04 4C 04 3D 04 38 04 3A 04 20 00 32 00 00 00 " +
                    "23 00 22 F1 4E 07 00 00 FF 01 00 00 40 00 A9 C3 " +
                    "42 07 00 00 50 4B 03 04 14 00 06 00 08 00 00 00 " +
                    "21 00 F0 F7 8A BB FD 00 00 00 E2 01 00 00 13 00 " +
                    "00 00 5B 43 6F 6E 74 65 6E 74 5F 54 79 70 65 73 " +
                    "5D 2E 78 6D 6C 94 91 CD 4A C4 30 10 C7 EF 82 EF " +
                    "10 E6 2A 6D AA 07 11 69 BA 07 AB 47 15 5D 1F 60 " +
                    "48 A6 6D D8 36 09 99 58 77 DF DE 74 3F 2E E2 0A " +
                    "1E 67 E6 FF F1 23 A9 57 DB 69 14 33 45 B6 DE 29 " +
                    "B8 2E 2B 10 E4 B4 37 D6 F5 0A 3E D6 4F C5 1D 08 " +
                    "4E E8 0C 8E DE 91 82 1D 31 AC 9A CB 8B 7A BD 0B " +
                    "C4 22 BB 1D 2B 18 52 0A F7 52 B2 1E 68 42 2E 7D " +
                    "20 97 2F 9D 8F 13 A6 3C C6 5E 06 D4 1B EC 49 DE " +
                    "54 D5 AD D4 DE 25 72 A9 48 4B 06 34 75 4B 1D 7E " +
                    "8E 49 3C 6E F3 FA 40 12 69 64 10 0F 07 E1 D2 A5 " +
                    "00 43 18 AD C6 94 49 E5 EC CC 8F 96 E2 D8 50 66 " +
                    "E7 5E C3 83 0D 7C 95 31 40 FE DA B0 5C CE 17 1C " +
                    "7D 2F F9 69 A2 35 24 5E 31 A6 67 9C 32 86 34 91 " +
                    "25 0F 18 28 6B CA BF 53 16 CC 89 0B DF 75 56 53 " +
                    "D9 46 7E 5F 7C 27 A8 73 E1 C6 7F B9 48 F3 7F B3 " +
                    "DB 6C 7B A3 F9 94 2E F7 3F D4 7C 03 00 00 FF FF " +
                    "03 00 50 4B 03 04 14 00 06 00 08 00 00 00 21 00 " +
                    "31 DD 5F 61 D2 00 00 00 8F 01 00 00 0B 00 00 00 " +
                    "5F 72 65 6C 73 2F 2E 72 65 6C 73 A4 90 C1 6A C3 " +
                    "30 0C 86 EF 83 BD 83 D1 BD 71 DA 43 19 A3 4E 6F " +
                    "85 5E 4B 07 BB 0A 5B 49 4C 63 CB 58 26 6D DF BE " +
                    "A6 30 58 46 6F 3B EA 17 FA 3E F1 EF F6 B7 30 A9 " +
                    "99 B2 78 8E 06 D6 4D 0B 8A A2 65 E7 E3 60 E0 EB " +
                    "7C 58 7D 80 92 82 D1 E1 C4 91 0C DC 49 60 DF BD " +
                    "BF ED 4E 34 61 A9 47 32 FA 24 AA 52 A2 18 18 4B " +
                    "49 9F 5A 8B 1D 29 A0 34 9C 28 D6 4D CF 39 60 A9 " +
                    "63 1E 74 42 7B C1 81 F4 A6 6D B7 3A FF 66 40 B7 " +
                    "60 AA A3 33 90 8F 6E 03 EA 7C 4F D5 FC 87 1D BC " +
                    "CD 2C DC 97 C6 72 D0 DC F7 DE BE A2 6A C7 D7 78 " +
                    "A2 B9 52 30 0F 54 0C B8 2C CF 30 D3 DC D4 E7 40 " +
                    "BF F6 AE FF E9 95 11 13 7D 57 FE 42 FC 4C AB F5 " +
                    "C7 AC 17 35 76 0F 00 00 00 FF FF 03 00 50 4B 03 " +
                    "04 14 00 06 00 08 00 00 00 21 00 54 76 AD 8A 97 " +
                    "02 00 00 FE 06 00 00 10 00 00 00 64 72 73 2F 73 " +
                    "68 61 70 65 78 6D 6C 2E 78 6D 6C AC 55 CD 8E D3 " +
                    "30 10 BE 23 F1 0E 96 EF BB F9 69 9B 2D 51 93 15 " +
                    "74 05 17 C4 56 5D F6 01 4C E2 B4 11 8E 1D D9 A6 " +
                    "3F 7B 42 E2 8A B4 8F C0 43 70 41 FC EC 33 A4 6F " +
                    "C4 D8 4E 52 60 81 C3 B6 3D A4 CE 4C 3C DF 37 DF " +
                    "CC D8 93 F3 4D C5 D0 8A 4A 55 0A 9E E0 E0 D4 C7 " +
                    "88 F2 4C E4 25 5F 24 F8 FA F5 F3 93 31 46 4A 13 " +
                    "9E 13 26 38 4D F0 96 2A 7C 9E 3E 7E 34 D9 E4 32 " +
                    "26 3C 5B 0A 89 20 04 57 31 18 12 BC D4 BA 8E 3D " +
                    "4F 65 4B 5A 11 75 2A 6A CA C1 5B 08 59 11 0D AF " +
                    "72 E1 E5 92 AC 21 78 C5 BC D0 F7 23 4F D5 92 92 " +
                    "5C 2D 29 D5 17 CE 83 53 1B 5B AF C5 94 32 F6 D4 " +
                    "42 38 53 21 45 E5 56 99 60 69 38 F1 0C 07 B3 B4 " +
                    "1B 60 71 59 14 E9 30 8A CE C2 51 EF 33 26 EB 96 " +
                    "62 9D 0E 9D D9 2C 3B 9B F1 07 C1 70 E0 FB BD CF " +
                    "6E B1 B1 F7 80 5A F4 20 E9 3E 78 6F B3 51 C2 C1 " +
                    "F8 5F C0 41 4B F6 1E 72 14 3C E9 F6 80 6F 8F DC " +
                    "E1 A9 1A 55 24 93 22 C1 18 69 BA D1 AC E4 6F 61 " +
                    "ED 80 F9 EA AA 9E C9 96 C4 AB D5 4C A2 32 4F F0 " +
                    "00 23 4E 2A 28 55 F3 69 F7 7E 77 DB 7C 6F EE 76 " +
                    "1F 9A CF CD 5D F3 6D F7 B1 F9 D1 7C 69 BE A2 10 " +
                    "7B FD 36 13 03 DE 6C C2 BF 46 54 36 36 89 37 85 " +
                    "AC DA 12 93 07 14 B8 22 25 07 BE 24 16 45 81 36 " +
                    "D0 62 D1 78 04 39 63 B4 4D F0 F8 2C 02 E1 0D 17 " +
                    "12 43 72 28 33 FE 21 F8 C1 88 32 F8 20 18 9D 05 " +
                    "11 7C 6C F8 39 26 E6 D3 5A 2A FD 82 8A 83 59 21 " +
                    "13 28 C1 92 66 DA 32 24 AB 97 4A 3B A8 0E A2 D5 " +
                    "C5 69 61 DA 4D E9 2D A3 86 04 E3 73 0A 09 D9 D6 " +
                    "7F B0 2E 50 30 48 39 B4 E8 76 66 E8 94 49 B4 22 " +
                    "2C C1 24 CB 28 D7 81 73 2D 49 4E 9D 79 E4 C3 AF " +
                    "D5 A3 DF 61 D5 B1 84 0C B3 A2 64 EC 68 DC 5A 02 " +
                    "66 9E EF 73 73 5A B5 78 B6 88 45 01 62 1E 0D DC " +
                    "FF 9F 30 0E 9C 76 88 36 73 C1 8F 07 5E 95 5C C8 " +
                    "BF 11 60 50 95 36 73 87 D7 35 89 6B 0D D3 25 7A " +
                    "F3 4C E4 5B 43 E9 0D FC C3 60 1E DA 27 70 42 EB " +
                    "4B 78 14 4C AC 13 9C B1 B2 C6 08 8E DE 9B 3F 6D " +
                    "52 B3 A9 80 EE 81 F9 71 87 73 82 B5 9B 2F A6 F4 " +
                    "95 21 78 28 15 C8 1C 26 F0 D0 28 36 08 E8 42 D8 " +
                    "02 AE 1D E6 28 52 9E CF 88 24 73 B0 33 62 EE 1F " +
                    "F9 EE 64 7E 0D F7 CF 0D 9C 04 41 DF F6 75 AB 77 " +
                    "27 B2 3D B9 14 58 ED 75 C0 4A 18 9B 0B A2 89 29 " +
                    "91 AD C5 EF 17 89 B5 39 6D D2 9F 00 00 00 FF FF " +
                    "03 00 50 4B 03 04 14 00 06 00 08 00 00 00 21 00 " +
                    "88 BE 95 E9 1C 01 00 00 90 01 00 00 0F 00 00 00 " +
                    "64 72 73 2F 64 6F 77 6E 72 65 76 2E 78 6D 6C 54 " +
                    "90 4D 4F C3 30 0C 86 EF 48 FC 87 C8 48 5C 10 4B " +
                    "3F D6 52 CA B2 69 20 A1 71 01 B1 AD 1C B8 85 26 " +
                    "5D 2B 9A A4 4A C2 DA FD 7B 3C 06 DA B8 F9 B5 FD " +
                    "D8 7E 3D 99 0D AA 25 5B 69 5D 63 34 83 70 14 00 " +
                    "91 BA 34 A2 D1 1B 06 C5 FA F1 3A 03 E2 3C D7 82 " +
                    "B7 46 4B 06 3B E9 60 36 3D 3F 9B F0 5C 98 5E 2F " +
                    "E5 76 E5 37 04 87 68 97 73 06 B5 F7 5D 4E A9 2B " +
                    "6B A9 B8 1B 99 4E 6A AC 55 C6 2A EE 51 DA 0D 15 " +
                    "96 F7 38 5C B5 34 0A 82 94 2A DE 68 DC 50 F3 4E " +
                    "3E D4 B2 FC 5C 7D 29 5C F2 AA DE 0A 73 9F BD 3F " +
                    "D3 AB A2 5F 2E D6 59 93 C4 19 63 97 17 C3 FC 0E " +
                    "88 97 83 3F 36 FF D2 4F 82 41 0C A4 5A EC 3E 6C " +
                    "23 96 DC 79 69 19 A0 1D 34 87 C6 60 8A 17 0F ED " +
                    "5C 97 B5 B1 FB B8 B2 46 11 6B 7A 06 63 20 A5 69 " +
                    "19 44 B0 D7 2F 55 E5 A4 47 22 1C C7 01 E2 58 FA " +
                    "4B 8D D3 F4 26 4A 80 EE 71 6F 0E 70 88 D4 0F 9D " +
                    "FC A7 D3 F0 16 5B 4F E9 30 8A B3 03 4D 4F 2F 41 " +
                    "71 7C E4 F4 1B 00 00 FF FF 03 00 50 4B 01 02 2D " +
                    "00 14 00 06 00 08 00 00 00 21 00 F0 F7 8A BB FD " +
                    "00 00 00 E2 01 00 00 13 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 00 00 00 00 5B 43 6F 6E 74 65 6E " +
                    "74 5F 54 79 70 65 73 5D 2E 78 6D 6C 50 4B 01 02 " +
                    "2D 00 14 00 06 00 08 00 00 00 21 00 31 DD 5F 61 " +
                    "D2 00 00 00 8F 01 00 00 0B 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 00 2E 01 00 00 5F 72 65 6C 73 2F " +
                    "2E 72 65 6C 73 50 4B 01 02 2D 00 14 00 06 00 08 " +
                    "00 00 00 21 00 54 76 AD 8A 97 02 00 00 FE 06 00 " +
                    "00 10 00 00 00 00 00 00 00 00 00 00 00 00 00 29 " +
                    "02 00 00 64 72 73 2F 73 68 61 70 65 78 6D 6C 2E " +
                    "78 6D 6C 50 4B 01 02 2D 00 14 00 06 00 08 00 00 " +
                    "00 21 00 88 BE 95 E9 1C 01 00 00 90 01 00 00 0F " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 00 EE 04 00 " +
                    "00 64 72 73 2F 64 6F 77 6E 72 65 76 2E 78 6D 6C " +
                    "50 4B 05 06 00 00 00 00 04 00 04 00 F5 00 00 00 " +
                    "37 06 00 00 00 00 00 00 10 F0 12 00 00 00 00 00 " +
                    "02 00 10 03 04 00 9A 00 05 00 D0 00 0C 00 DA 00 " +
                    "00 00 11 F0 00 00 00 00 5D 00 1A 00 15 00 12 00 " +
                    "02 00 02 00 11 60 00 00 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 00 EC 00 E3 07 0F 00 04 F0 DB 07 " +
                    "00 00 12 00 0A F0 08 00 00 00 03 04 00 00 00 0A " +
                    "00 00 83 00 0B F0 50 00 00 00 BF 00 18 00 1F 00 " +
                    "81 01 4F 81 BD 00 BF 01 10 00 10 00 C0 01 38 5D " +
                    "8A 00 CB 01 38 63 00 00 FF 01 08 00 08 00 80 C3 " +
                    "20 00 00 00 BF 03 00 00 02 00 1F 04 40 04 4F 04 " +
                    "3C 04 3E 04 43 04 33 04 3E 04 3B 04 4C 04 3D 04 " +
                    "38 04 3A 04 20 00 33 00 00 00 23 00 22 F1 49 07 " +
                    "00 00 FF 01 00 00 40 00 A9 C3 3D 07 00 00 50 4B " +
                    "03 04 14 00 06 00 08 00 00 00 21 00 F0 F7 8A BB " +
                    "FD 00 00 00 E2 01 00 00 13 00 00 00 5B 43 6F 6E " +
                    "74 65 6E 74 5F 54 79 70 65 73 5D 2E 78 6D 6C 94 " +
                    "91 CD 4A C4 30 10 C7 EF 82 EF 10 E6 2A 6D AA 07 " +
                    "11 69 BA 07 AB 47 15 5D 1F 60 48 A6 6D D8 36 09 " +
                    "99 58 77 DF DE 74 3F 2E E2 0A 1E 67 E6 FF F1 23 " +
                    "A9 57 DB 69 14 33 45 B6 DE 29 B8 2E 2B 10 E4 B4 " +
                    "37 D6 F5 0A 3E D6 4F C5 1D 08 4E E8 0C 8E DE 91 " +
                    "82 1D 31 AC 9A CB 8B 7A BD 0B C4 22 BB 1D 2B 18 " +
                    "52 0A F7 52 B2 1E 68 42 2E 7D 20 97 2F 9D 8F 13 " +
                    "A6 3C C6 5E 06 D4 1B EC 49 DE 54 D5 AD D4 DE 25 " +
                    "72 A9 48 4B 06 34 75 4B 1D 7E 8E 49 3C 6E F3 FA " +
                    "40 12 69 64 10 0F 07 E1 D2 A5 00 43 18 AD C6 94 " +
                    "49 E5 EC CC 8F 96 E2 D8 50 66 E7 5E C3 83 0D 7C " +
                    "95 31 40 FE DA B0 5C CE 17 1C 7D 2F F9 69 A2 35 " +
                    "24 5E 31 A6 67 9C 32 86 34 91 25 0F 18 28 6B CA " +
                    "BF 53 16 CC 89 0B DF 75 56 53 D9 46 7E 5F 7C 27 " +
                    "A8 73 E1 C6 7F B9 48 F3 7F B3 DB 6C 7B A3 F9 94 " +
                    "2E F7 3F D4 7C 03 00 00 FF FF 03 00 50 4B 03 04 " +
                    "14 00 06 00 08 00 00 00 21 00 31 DD 5F 61 D2 00 " +
                    "00 00 8F 01 00 00 0B 00 00 00 5F 72 65 6C 73 2F " +
                    "2E 72 65 6C 73 A4 90 C1 6A C3 30 0C 86 EF 83 BD " +
                    "83 D1 BD 71 DA 43 19 A3 4E 6F 85 5E 4B 07 BB 0A " +
                    "5B 49 4C 63 CB 58 26 6D DF BE A6 30 58 46 6F 3B " +
                    "EA 17 FA 3E F1 EF F6 B7 30 A9 99 B2 78 8E 06 D6 " +
                    "4D 0B 8A A2 65 E7 E3 60 E0 EB 7C 58 7D 80 92 82 " +
                    "D1 E1 C4 91 0C DC 49 60 DF BD BF ED 4E 34 61 A9 " +
                    "47 32 FA 24 AA 52 A2 18 18 4B 49 9F 5A 8B 1D 29 " +
                    "A0 34 9C 28 D6 4D CF 39 60 A9 63 1E 74 42 7B C1 " +
                    "81 F4 A6 6D B7 3A FF 66 40 B7 60 AA A3 33 90 8F " +
                    "6E 03 EA 7C 4F D5 FC 87 1D BC CD 2C DC 97 C6 72 " +
                    "D0 DC F7 DE BE A2 6A C7 D7 78 A2 B9 52 30 0F 54 " +
                    "0C B8 2C CF 30 D3 DC D4 E7 40 BF F6 AE FF E9 95 " +
                    "11 13 7D 57 FE 42 FC 4C AB F5 C7 AC 17 35 76 0F " +
                    "00 00 00 FF FF 03 00 50 4B 03 04 14 00 06 00 08 " +
                    "00 00 00 21 00 B1 89 CE 0F 93 02 00 00 FC 06 00 " +
                    "00 10 00 00 00 64 72 73 2F 73 68 61 70 65 78 6D " +
                    "6C 2E 78 6D 6C AC 55 4B 6E DB 30 10 DD 17 E8 1D " +
                    "08 EE 13 7D 1C D9 8A 60 29 68 1D B4 9B A2 31 9C " +
                    "E6 00 AC 44 D9 42 29 52 20 59 7F B2 2A D0 6D 81 " +
                    "1E A1 87 E8 A6 E8 27 67 90 6F D4 21 29 B9 8D FB " +
                    "59 C4 F6 C2 1E CF 70 E6 BD F9 91 E3 8B 75 CD D0 " +
                    "92 4A 55 09 9E E2 E0 D4 C7 88 F2 5C 14 15 9F A7 " +
                    "F8 E6 D5 B3 93 18 23 A5 09 2F 08 13 9C A6 78 43 " +
                    "15 BE C8 1E 3F 1A AF 0B 99 10 9E 2F 84 44 10 82 " +
                    "AB 04 14 29 5E 68 DD 24 9E A7 F2 05 AD 89 3A 15 " +
                    "0D E5 60 2D 85 AC 89 86 BF 72 EE 15 92 AC 20 78 " +
                    "CD BC D0 F7 87 9E 6A 24 25 85 5A 50 AA 2F 9D 05 " +
                    "67 36 B6 5E 89 09 65 EC 89 85 70 AA 52 8A DA 49 " +
                    "B9 60 D9 60 EC 19 0E 46 B4 0E 20 5C 95 65 76 1E " +
                    "85 D1 CE 62 14 D6 28 C5 2A EB D4 46 EC 75 C6 3E " +
                    "1A 02 11 E7 01 26 EB 61 03 FF 42 D3 62 87 D0 07 " +
                    "D9 47 0D 21 CA BF 70 83 8E E9 3E 70 10 0E E2 DE " +
                    "E7 1E 72 8F A7 1A 54 93 5C 8A 14 63 A4 E9 5A B3 " +
                    "8A BF 01 D9 91 E1 CB EB 66 2A 3B 62 2F 97 53 89 " +
                    "AA 22 C5 67 18 71 52 43 9F DA 4F DB 77 DB 8F ED " +
                    "F7 F6 6E FB BE FD DC DE B5 DF B6 1F DA 1F ED 97 " +
                    "F6 2B 1A 60 6F E7 66 62 C0 3F 9B F0 EF 11 95 8D " +
                    "4D 92 75 29 EB AE BF E4 01 DD AD 49 C5 81 2F 49 " +
                    "44 59 A2 35 CC 57 3C 88 07 61 84 D1 06 64 3F 8C " +
                    "47 BE 6F C8 90 04 B2 43 B9 39 70 16 47 E7 A0 44 " +
                    "B9 39 11 8D 82 21 9C 36 04 1D 15 73 B4 91 4A 3F " +
                    "A7 E2 60 5A C8 04 4A B1 A4 B9 B6 14 C9 F2 85 D2 " +
                    "0E AA 87 E8 0A E3 8A 61 86 4D E9 0D A3 86 04 E3 " +
                    "33 0A 19 D9 C1 7F 70 61 A0 63 90 72 68 D1 ED C6 " +
                    "D0 09 93 68 49 58 8A 49 9E 53 AE 03 67 5A 90 82 " +
                    "3A 75 E4 C3 A7 AB C7 CE C3 56 C7 12 32 CC CA 8A " +
                    "B1 A3 71 EB 08 98 6D FE 93 9B AB 55 87 67 9B 58 " +
                    "96 50 CC A3 81 FB FF 2B 8C 03 A7 3D A2 CD 5C F0 " +
                    "E3 81 D7 15 17 F2 6F 04 18 74 A5 CB DC E1 F5 43 " +
                    "E2 46 C3 4C 89 5E 3F 15 C5 C6 50 7A 0D BF B0 99 " +
                    "87 CE 09 DC CF FA 0A BE 4A 26 56 29 CE 59 D5 60 " +
                    "04 17 EF ED BE 4E 6A 36 11 30 3D B0 3F EE 6A 4E " +
                    "B1 76 FB C5 94 BE 36 04 0F A5 02 99 C3 06 1E 1A " +
                    "C5 06 81 BA 10 36 87 47 87 39 8A 94 17 53 22 C9 " +
                    "0C F4 8C 98 D7 47 BE 3D 99 DD C0 EB 73 0B 37 41 " +
                    "B0 1B FB A6 AB 77 5F 64 7B 75 29 D0 DA C7 80 55 " +
                    "B0 36 97 44 13 D3 22 DB 8B FB CF 88 D5 B9 DA 64 " +
                    "3F 01 00 00 FF FF 03 00 50 4B 03 04 14 00 06 00 " +
                    "08 00 00 00 21 00 48 FA 47 DC 1B 01 00 00 8D 01 " +
                    "00 00 0F 00 00 00 64 72 73 2F 64 6F 77 6E 72 65 " +
                    "76 2E 78 6D 6C 4C 90 CB 4E C3 30 10 45 F7 48 FC " +
                    "83 35 48 6C 10 75 92 92 12 42 9D AA 20 A1 B2 01 " +
                    "91 36 2C D8 99 C4 79 88 D8 8E 6C D3 A4 7F CF A4 " +
                    "A5 6A 77 BE BE 73 EE 3C E6 8B 41 B6 64 2B 8C 6D " +
                    "B4 62 E0 4F 3C 20 42 E5 BA 68 54 C5 20 DB BC DC " +
                    "46 40 AC E3 AA E0 AD 56 82 C1 4E 58 58 24 97 17 " +
                    "73 1E 17 BA 57 A9 D8 AE 5D 45 30 44 D9 98 33 A8 " +
                    "9D EB 62 4A 6D 5E 0B C9 ED 44 77 42 A1 57 6A 23 " +
                    "B9 43 69 2A 5A 18 DE 63 B8 6C 69 E0 79 33 2A 79 " +
                    "A3 B0 43 CD 3B F1 5C 8B FC 67 FD 2B B1 C9 87 FC " +
                    "CC F4 53 F4 F5 46 6F B2 3E 5D 6D A2 26 9C 46 8C " +
                    "5D 5F 0D CB 47 20 4E 0C EE 54 FC 4F BF 16 0C EE " +
                    "80 94 AB DD B7 69 8A 94 5B 27 0C 03 5C 07 97 C3 " +
                    "C5 20 C1 89 87 76 A9 F2 5A 9B F1 5D 1A 2D 89 D1 " +
                    "3D 83 10 48 AE 5B 06 53 18 F5 7B 59 5A E1 18 DC " +
                    "CF 70 BE BD 73 FC 79 08 83 10 E8 C8 3A 7D 20 7D " +
                    "44 F6 28 46 9C A1 7E 30 8D B0 74 B4 8E 6C 80 71 " +
                    "07 9A 9E 8F 81 E2 74 C5 E4 0F 00 00 FF FF 03 00 " +
                    "50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 21 00 " +
                    "F0 F7 8A BB FD 00 00 00 E2 01 00 00 13 00 00 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 00 00 5B 43 " +
                    "6F 6E 74 65 6E 74 5F 54 79 70 65 73 5D 2E 78 6D " +
                    "6C 50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 21 " +
                    "00 31 DD 5F 61 D2 00 00 00 8F 01 00 00 0B 00 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 2E 01 00 00 5F " +
                    "72 65 6C 73 2F 2E 72 65 6C 73 50 4B 01 02 2D 00 " +
                    "14 00 06 00 08 00 00 00 21 00 B1 89 CE 0F 93 02 " +
                    "00 00 FC 06 00 00 10 00 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 29 02 00 00 64 72 73 2F 73 68 61 70 " +
                    "65 78 6D 6C 2E 78 6D 6C 50 4B 01 02 2D 00 14 00 " +
                    "06 00 08 00 00 00 21 00 48 FA 47 DC 1B 01 00 00 " +
                    "8D 01 00 00 0F 00 00 00 00 00 00 00 00 00 00 00 " +
                    "00 00 EA 04 00 00 64 72 73 2F 64 6F 77 6E 72 65 " +
                    "76 2E 78 6D 6C 50 4B 05 06 00 00 00 00 04 00 04 " +
                    "00 F5 00 00 00 32 06 00 00 00 00 00 00 10 F0 12 " +
                    "00 00 00 00 00 03 00 10 00 05 00 66 00 05 00 D0 " +
                    "01 0D 00 A6 00 00 00 11 F0 00 00 00 00 5D 00 1A " +
                    "00 15 00 12 00 02 00 03 00 11 60 00 00 00 00 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 EC 00 E8 07 0F " +
                    "00 04 F0 E0 07 00 00 12 00 0A F0 08 00 00 00 04 " +
                    "04 00 00 00 0A 00 00 83 00 0B F0 50 00 00 00 BF " +
                    "00 18 00 1F 00 81 01 4F 81 BD 00 BF 01 10 00 10 " +
                    "00 C0 01 38 5D 8A 00 CB 01 38 63 00 00 FF 01 08 " +
                    "00 08 00 80 C3 20 00 00 00 BF 03 00 00 02 00 1F " +
                    "04 40 04 4F 04 3C 04 3E 04 43 04 33 04 3E 04 3B " +
                    "04 4C 04 3D 04 38 04 3A 04 20 00 34 00 00 00 23 " +
                    "00 22 F1 4E 07 00 00 FF 01 00 00 40 00 A9 C3 42 " +
                    "07 00 00 50 4B 03 04 14 00 06 00 08 00 00 00 21 " +
                    "00 F0 F7 8A BB FD 00 00 00 E2 01 00 00 13 00 00 " +
                    "00 5B 43 6F 6E 74 65 6E 74 5F 54 79 70 65 73 5D " +
                    "2E 78 6D 6C 94 91 CD 4A C4 30 10 C7 EF 82 EF 10 " +
                    "E6 2A 6D AA 07 11 69 BA 07 AB 47 15 5D 1F 60 48 " +
                    "A6 6D D8 36 09 99 58 77 DF DE 74 3F 2E E2 0A 1E " +
                    "67 E6 FF F1 23 A9 57 DB 69 14 33 45 B6 DE 29 B8 " +
                    "2E 2B 10 E4 B4 37 D6 F5 0A 3E D6 4F C5 1D 08 4E " +
                    "E8 0C 8E DE 91 82 1D 31 AC 9A CB 8B 7A BD 0B C4 " +
                    "22 BB 1D 2B 18 52 0A F7 52 B2 1E 68 42 2E 7D 20 " +
                    "97 2F 9D 8F 13 A6 3C C6 5E 06 D4 1B EC 49 DE 54 " +
                    "D5 AD D4 DE 25 72 A9 48 4B 06 34 75 4B 1D 7E 8E " +
                    "49 3C 6E F3 FA 40 12 69 64 10 0F 07 E1 D2 A5 00 " +
                    "43 18 AD C6 94 49 E5 EC CC 8F 96 E2 D8 50 66 E7 " +
                    "5E C3 83 0D 7C 95 31 40 FE DA B0 5C CE 17 1C 7D " +
                    "2F F9 69 A2 35 24 5E 31 A6 67 9C 32 86 34 91 25 " +
                    "0F 18 28 6B CA BF 53 16 CC 89 0B DF 75 56 53 D9 " +
                    "46 7E 5F 7C 27 A8 73 E1 C6 7F B9 48 F3 7F B3 DB " +
                    "6C 7B A3 F9 94 2E F7 3F D4 7C 03 00 00 FF FF 03 " +
                    "00 50 4B 03 04 14 00 06 00 08 00 00 00 21 00 31 " +
                    "DD 5F 61 D2 00 00 00 8F 01 00 00 0B 00 00 00 5F " +
                    "72 65 6C 73 2F 2E 72 65 6C 73 A4 90 C1 6A C3 30 " +
                    "0C 86 EF 83 BD 83 D1 BD 71 DA 43 19 A3 4E 6F 85 " +
                    "5E 4B 07 BB 0A 5B 49 4C 63 CB 58 26 6D DF BE A6 " +
                    "30 58 46 6F 3B EA 17 FA 3E F1 EF F6 B7 30 A9 99 " +
                    "B2 78 8E 06 D6 4D 0B 8A A2 65 E7 E3 60 E0 EB 7C " +
                    "58 7D 80 92 82 D1 E1 C4 91 0C DC 49 60 DF BD BF " +
                    "ED 4E 34 61 A9 47 32 FA 24 AA 52 A2 18 18 4B 49 " +
                    "9F 5A 8B 1D 29 A0 34 9C 28 D6 4D CF 39 60 A9 63 " +
                    "1E 74 42 7B C1 81 F4 A6 6D B7 3A FF 66 40 B7 60 " +
                    "AA A3 33 90 8F 6E 03 EA 7C 4F D5 FC 87 1D BC CD " +
                    "2C DC 97 C6 72 D0 DC F7 DE BE A2 6A C7 D7 78 A2 " +
                    "B9 52 30 0F 54 0C B8 2C CF 30 D3 DC D4 E7 40 BF " +
                    "F6 AE FF E9 95 11 13 7D 57 FE 42 FC 4C AB F5 C7 " +
                    "AC 17 35 76 0F 00 00 00 FF FF 03 00 50 4B 03 04 " +
                    "14 00 06 00 08 00 00 00 21 00 CC C2 D3 C8 98 02 " +
                    "00 00 FD 06 00 00 10 00 00 00 64 72 73 2F 73 68 " +
                    "61 70 65 78 6D 6C 2E 78 6D 6C AC 55 4B 6E DB 30 " +
                    "10 DD 17 E8 1D 08 EE 13 49 8E E5 C8 82 A5 A0 75 " +
                    "D0 6E 8A C6 70 9A 03 B0 12 65 0B A5 48 81 64 FD " +
                    "C9 AA 40 B7 05 7A 84 1E A2 9B A2 9F 9C 41 BE 51 " +
                    "87 A4 24 37 FD 2D 62 7B 61 D3 33 E4 BC 37 6F 66 " +
                    "C8 C9 C5 A6 62 68 45 A5 2A 05 4F 70 70 EA 63 44 " +
                    "79 26 F2 92 2F 12 7C F3 EA D9 49 84 91 D2 84 E7 " +
                    "84 09 4E 13 BC A5 0A 5F A4 8F 1F 4D 36 B9 8C 09 " +
                    "CF 96 42 22 08 C1 55 0C 86 04 2F B5 AE 63 CF 53 " +
                    "D9 92 56 44 9D 8A 9A 72 F0 16 42 56 44 C3 5F B9 " +
                    "F0 72 49 D6 10 BC 62 DE C0 F7 47 9E AA 25 25 B9 " +
                    "5A 52 AA 2F 9D 07 A7 36 B6 5E 8B 29 65 EC 89 85 " +
                    "70 A6 42 8A CA AD 32 C1 D2 B3 89 67 38 98 A5 3D " +
                    "00 8B AB A2 48 83 51 30 1E 84 BD CF 98 AC 5B 8A " +
                    "75 3A 72 66 B3 EC 6C C6 7F 16 05 BE DF BB EC 09 " +
                    "1B 7A 8F A7 45 8F 91 EE 63 F7 36 73 64 38 88 46 " +
                    "FF C2 0D 86 7D F4 7B C0 51 78 DE 1D 01 4E 7B E0 " +
                    "0E 4E D5 A8 22 99 14 09 C6 48 D3 8D 66 25 7F 03 " +
                    "6B 17 83 AF AE EB 99 6C 39 BC 5C CD 24 2A F3 04 " +
                    "87 18 71 52 41 A1 9A 4F BB 77 BB 8F CD F7 E6 6E " +
                    "F7 BE F9 DC DC 35 DF 76 1F 9A 1F CD 97 E6 2B 1A " +
                    "62 AF 3F 66 62 C0 3F 9B EF AF 11 95 8D 4D E2 4D " +
                    "21 AB B6 C0 E4 01 E5 AD 48 C9 81 2F 89 45 51 A0 " +
                    "0D 34 D8 78 EC 43 CE 18 6D 61 1D 44 01 08 6F C8 " +
                    "90 18 B2 43 99 D9 30 8C C2 31 18 51 66 76 84 E7 " +
                    "01 88 6A 76 78 8E 8A D9 5A 4B A5 9F 53 71 30 2D " +
                    "64 02 25 58 D2 4C 5B 8A 64 F5 42 69 07 D5 41 B4 " +
                    "C2 38 31 4C B7 29 BD 65 D4 90 60 7C 4E 21 23 DB " +
                    "F9 0F 16 06 2A 06 29 0F 2C BA 1D 19 3A 65 12 AD " +
                    "08 4B 30 C9 32 CA 75 E0 5C 4B 92 53 67 0E 7D F8 " +
                    "B4 7A F4 27 AC 3A 96 90 61 56 94 8C 1D 8D 5B 4B " +
                    "C0 8C F3 9F DC 9C 56 2D 9E 2D 62 51 80 98 47 03 " +
                    "F7 FF 27 8C 03 A7 1D A2 CD 5C F0 E3 81 57 25 17 " +
                    "F2 6F 04 18 54 A5 CD DC E1 75 4D E2 5A C3 74 89 " +
                    "DE 3C 15 F9 D6 50 7A 0D BF 30 99 87 F6 09 5C D0 " +
                    "FA 0A BE 0A 26 D6 09 CE 58 59 63 04 37 EF ED EF " +
                    "36 A9 D9 54 40 F7 C0 FC B8 BB 39 C1 DA CD 17 53 " +
                    "FA DA 10 3C 94 0A 64 0E 13 78 68 14 1B 04 74 21 " +
                    "6C 01 AF 0E 73 14 29 CF 67 44 92 39 D8 19 31 CF " +
                    "8F 7C 7B 32 BF 81 E7 E7 D6 DC 15 7D DB D7 AD DE " +
                    "9D C8 F6 EA 52 60 B5 AF 01 2B 61 6C 2E 89 26 A6 " +
                    "44 B6 16 F7 DF 11 6B 73 DA A4 3F 01 00 00 FF FF " +
                    "03 00 50 4B 03 04 14 00 06 00 08 00 00 00 21 00 " +
                    "A9 7A 63 39 1B 01 00 00 8E 01 00 00 0F 00 00 00 " +
                    "64 72 73 2F 64 6F 77 6E 72 65 76 2E 78 6D 6C 4C " +
                    "90 CD 4E C3 30 10 84 EF 48 BC 83 B5 48 5C 10 75 " +
                    "D2 36 C1 84 BA 55 41 42 ED 05 44 DA 70 E0 66 12 " +
                    "E7 47 C4 76 65 9B 26 7D 7B 36 6A 51 7B F3 78 F7 " +
                    "DB D9 D9 D9 A2 57 2D D9 4B EB 1A A3 39 84 A3 00 " +
                    "88 D4 B9 29 1A 5D 71 C8 B6 AF F7 0C 88 F3 42 17 " +
                    "A2 35 5A 72 38 48 07 8B F9 F5 D5 4C 24 85 E9 74 " +
                    "2A F7 1B 5F 11 1C A2 5D 22 38 D4 DE EF 12 4A 5D " +
                    "5E 4B 25 DC C8 EC A4 C6 5A 69 AC 12 1E A5 AD 68 " +
                    "61 45 87 C3 55 4B C7 41 10 53 25 1A 8D 0E B5 D8 " +
                    "C9 97 5A E6 3F 9B 5F 85 26 1F EA 33 33 CF EC EB " +
                    "8D DE 65 5D BA DA B2 26 9A 30 CE 6F 6F FA E5 13 " +
                    "10 2F 7B 7F 6E 3E D1 EB 82 43 04 A4 5C 1D BE 6D " +
                    "53 A4 C2 79 69 39 60 1C 0C 87 C1 60 8E 1B F7 ED " +
                    "52 E7 B5 B1 C3 BB B4 46 11 6B 3A 0E 31 90 DC B4 " +
                    "1C 26 30 E8 F7 B2 74 D2 A3 62 61 80 34 56 FE 7F " +
                    "C2 38 7C 1C 47 40 07 DA 9B 23 1B 4E 4F 30 5A 5F " +
                    "C0 2C 7A C0 CE 4B 78 3A 66 F1 11 A6 97 7B A0 38 " +
                    "9F 71 FE 07 00 00 FF FF 03 00 50 4B 01 02 2D 00 " +
                    "14 00 06 00 08 00 00 00 21 00 F0 F7 8A BB FD 00 " +
                    "00 00 E2 01 00 00 13 00 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 00 00 00 5B 43 6F 6E 74 65 6E 74 " +
                    "5F 54 79 70 65 73 5D 2E 78 6D 6C 50 4B 01 02 2D " +
                    "00 14 00 06 00 08 00 00 00 21 00 31 DD 5F 61 D2 " +
                    "00 00 00 8F 01 00 00 0B 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 2E 01 00 00 5F 72 65 6C 73 2F 2E " +
                    "72 65 6C 73 50 4B 01 02 2D 00 14 00 06 00 08 00 " +
                    "00 00 21 00 CC C2 D3 C8 98 02 00 00 FD 06 00 00 " +
                    "10 00 00 00 00 00 00 00 00 00 00 00 00 00 29 02 " +
                    "00 00 64 72 73 2F 73 68 61 70 65 78 6D 6C 2E 78 " +
                    "6D 6C 50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 " +
                    "21 00 A9 7A 63 39 1B 01 00 00 8E 01 00 00 0F 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 EF 04 00 00 " +
                    "64 72 73 2F 64 6F 77 6E 72 65 76 2E 78 6D 6C 50 " +
                    "4B 05 06 00 00 00 00 04 00 04 00 F5 00 00 00 37 " +
                    "06 00 00 00 00 00 00 10 F0 12 00 00 00 00 00 03 " +
                    "00 10 01 06 00 33 00 05 00 D0 02 0E 00 73 00 00 " +
                    "00 11 F0 00 00 00 00 5D 00 1A 00 15 00 12 00 02 " +
                    "00 04 00 11 60 00 00 00 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 3C 00 E5 07 0F 00 04 F0 DD 07 00 " +
                    "00 12 00 0A F0 08 00 00 00 05 04 00 00 00 0A 00 " +
                    "00 83 00 0B F0 50 00 00 00 BF 00 18 00 1F 00 81 " +
                    "01 4F 81 BD 00 BF 01 10 00 10 00 C0 01 38 5D 8A " +
                    "00 CB 01 38 63 00 00 FF 01 08 00 08 00 80 C3 20 " +
                    "00 00 00 BF 03 00 00 02 00 1F 04 40 04 4F 04 3C " +
                    "04 3E 04 43 04 33 04 3E 04 3B 04 4C 04 3D 04 38 " +
                    "04 3A 04 20 00 35 00 00 00 23 00 22 F1 4B 07 00 " +
                    "00 FF 01 00 00 40 00 A9 C3 3F 07 00 00 50 4B 03 " +
                    "04 14 00 06 00 08 00 00 00 21 00 F0 F7 8A BB FD " +
                    "00 00 00 E2 01 00 00 13 00 00 00 5B 43 6F 6E 74 " +
                    "65 6E 74 5F 54 79 70 65 73 5D 2E 78 6D 6C 94 91 " +
                    "CD 4A C4 30 10 C7 EF 82 EF 10 E6 2A 6D AA 07 11 " +
                    "69 BA 07 AB 47 15 5D 1F 60 48 A6 6D D8 36 09 99 " +
                    "58 77 DF DE 74 3F 2E E2 0A 1E 67 E6 FF F1 23 A9 " +
                    "57 DB 69 14 33 45 B6 DE 29 B8 2E 2B 10 E4 B4 37 " +
                    "D6 F5 0A 3E D6 4F C5 1D 08 4E E8 0C 8E DE 91 82 " +
                    "1D 31 AC 9A CB 8B 7A BD 0B C4 22 BB 1D 2B 18 52 " +
                    "0A F7 52 B2 1E 68 42 2E 7D 20 97 2F 9D 8F 13 A6 " +
                    "3C C6 5E 06 D4 1B EC 49 DE 54 D5 AD D4 DE 25 72 " +
                    "A9 48 4B 06 34 75 4B 1D 7E 8E 49 3C 6E F3 FA 40 " +
                    "12 69 64 10 0F 07 E1 D2 A5 00 43 18 AD C6 94 49 " +
                    "E5 EC CC 8F 96 E2 D8 50 66 E7 5E C3 83 0D 7C 95 " +
                    "31 40 FE DA B0 5C CE 17 1C 7D 2F F9 69 A2 35 24 " +
                    "5E 31 A6 67 9C 32 86 34 91 25 0F 18 28 6B CA BF " +
                    "53 16 CC 89 0B DF 75 56 53 D9 46 7E 5F 7C 27 A8 " +
                    "73 E1 C6 7F B9 48 F3 7F B3 DB 6C 7B A3 F9 94 2E " +
                    "F7 3F D4 7C 03 00 00 FF FF 03 00 50 4B 03 04 14 " +
                    "00 06 00 08 00 00 00 21 00 31 DD 5F 61 D2 00 00 " +
                    "00 8F 01 00 00 0B 00 00 00 5F 72 65 6C 73 2F 2E " +
                    "72 65 6C 73 A4 90 C1 6A C3 30 0C 86 EF 83 BD 83 " +
                    "D1 BD 71 DA 43 19 A3 4E 6F 85 5E 4B 07 BB 0A 5B " +
                    "49 4C 63 CB 58 26 6D DF BE A6 30 58 46 6F 3B EA " +
                    "17 FA 3E F1 EF F6 B7 30 A9 99 B2 78 8E 06 D6 4D " +
                    "0B 8A A2 65 E7 E3 60 E0 EB 7C 58 7D 80 92 82 D1 " +
                    "E1 C4 91 0C DC 49 60 DF BD BF ED 4E 34 61 A9 47 " +
                    "32 FA 24 AA 52 A2 18 18 4B 49 9F 5A 8B 1D 29 A0 " +
                    "34 9C 28 D6 4D CF 39 60 A9 63 1E 74 42 7B C1 81 " +
                    "F4 A6 6D B7 3A FF 66 40 B7 60 AA A3 33 90 8F 6E " +
                    "03 EA 7C 4F D5 FC 87 1D BC CD 2C DC 97 C6 72 D0 " +
                    "DC F7 DE BE A2 6A C7 D7 78 A2 B9 52 30 0F 54 0C " +
                    "B8 2C CF 30 D3 DC D4 E7 40 BF F6 AE FF E9 95 11 " +
                    "13 7D 57 FE 42 FC 4C AB F5 C7 AC 17 35 76 0F 00 " +
                    "00 00 FF FF 03 00 50 4B 03 04 14 00 06 00 08 00 " +
                    "00 00 21 00 D1 22 B3 C1 96 02 00 00 F9 06 00 00 " +
                    "10 00 00 00 64 72 73 2F 73 68 61 70 65 78 6D 6C " +
                    "2E 78 6D 6C AC 55 4B 6E DB 30 10 DD 17 E8 1D 08 " +
                    "EE 13 49 96 E5 B8 82 A5 A0 75 D0 6E 8A C6 70 9A " +
                    "03 B0 12 65 0B A5 48 81 64 FD C9 AA 40 B7 05 72 " +
                    "84 1E A2 9B A2 9F 9C 41 BE 51 87 A4 24 B7 E9 67 " +
                    "11 7B 63 53 33 9C 79 8F 6F 66 C8 C9 F9 A6 62 68 " +
                    "45 A5 2A 05 4F 70 70 EA 63 44 79 26 F2 92 2F 12 " +
                    "7C FD FA F9 C9 18 23 A5 09 CF 09 13 9C 26 78 4B " +
                    "15 3E 4F 1F 3F 9A 6C 72 19 13 9E 2D 85 44 90 82 " +
                    "AB 18 0C 09 5E 6A 5D C7 9E A7 B2 25 AD 88 3A 15 " +
                    "35 E5 E0 2D 84 AC 88 86 4F B9 F0 72 49 D6 90 BC " +
                    "62 DE C0 F7 47 9E AA 25 25 B9 5A 52 AA 2F 9C 07 " +
                    "A7 36 B7 5E 8B 29 65 EC A9 85 70 A6 42 8A CA AD " +
                    "32 C1 D2 70 E2 19 0E 66 69 03 60 71 59 14 69 18 " +
                    "0C C3 41 D4 FB 8C C9 BA A5 58 A7 67 CE 6C 96 9D " +
                    "CD F8 FD DE 6C 77 DB B4 7B 2C 2D FA FC E9 3E 6F " +
                    "6F 33 21 D1 38 F0 FF 85 19 B4 31 F7 41 87 67 A3 " +
                    "2E 04 5C 7B E0 0E 4E D5 A8 22 99 14 09 C6 48 D3 " +
                    "8D 66 25 7F 0B 6B 87 CB 57 57 F5 4C B6 1C 5E AD " +
                    "66 12 95 79 82 47 18 71 52 41 91 9A 4F BB F7 BB " +
                    "DB E6 7B 73 B7 FB D0 7C 6E EE 9A 6F BB 8F CD 8F " +
                    "E6 4B F3 15 45 D8 EB C3 4C 0E F8 B2 E7 FD 35 A3 " +
                    "B2 B9 49 BC 29 64 D5 16 97 3C A0 B4 15 29 39 F0 " +
                    "25 B1 28 0A B4 49 F0 00 2A 13 0C 22 8C B6 D0 68 " +
                    "61 18 46 BE 6F C8 90 18 4E 87 32 D8 10 0C C7 D1 " +
                    "13 30 A2 CC EC 88 CE 02 50 C8 EC F0 1C 15 B3 B5 " +
                    "96 4A BF A0 E2 60 5A C8 24 4A B0 A4 99 B6 14 C9 " +
                    "EA A5 D2 0E AA 83 68 85 71 62 98 4E 53 7A CB A8 " +
                    "21 C1 F8 9C C2 89 6C D7 3F 58 18 A8 98 D1 C4 A2 " +
                    "DB 71 A1 53 26 D1 8A B0 04 93 2C A3 5C 07 CE B5 " +
                    "24 39 75 66 D0 CB 29 06 7A F4 11 56 1D 4B C8 30 " +
                    "2B 4A C6 8E C6 AD 25 60 46 F9 4F 6E 4E AB 16 CF " +
                    "16 B1 28 40 CC A3 81 FB FF 13 C6 81 D3 0E D1 9E " +
                    "5C F0 E3 81 57 25 17 F2 6F 04 18 54 A5 3D B9 C3 " +
                    "EB 9A C4 B5 86 E9 12 BD 79 26 F2 AD A1 F4 06 FE " +
                    "61 32 0F ED 13 B8 9C F5 25 FC 14 4C AC 13 9C B1 " +
                    "B2 C6 08 6E DD 9B FB 36 A9 D9 54 40 F7 C0 FC B8 " +
                    "7B 39 C1 DA CD 17 53 FA CA 10 3C 94 0A 9C 1C 26 " +
                    "F0 D0 2C 36 09 E8 42 D8 02 5E 1C D6 5E 01 3C 9F " +
                    "11 49 E6 60 67 C4 3C 3D F2 DD C9 FC 1A 9E 9E 1B " +
                    "B8 09 82 BE ED EB 56 EF 4E 64 7B 75 29 B0 DA 97 " +
                    "80 95 30 36 17 44 13 53 22 5B 8B DF DF 10 6B 73 " +
                    "DA A4 3F 01 00 00 FF FF 03 00 50 4B 03 04 14 00 " +
                    "06 00 08 00 00 00 21 00 EA 0B 8D A0 1A 01 00 00 " +
                    "8A 01 00 00 0F 00 00 00 64 72 73 2F 64 6F 77 6E " +
                    "72 65 76 2E 78 6D 6C 4C 90 5F 4F C2 30 14 C5 DF " +
                    "4D FC 0E CD 35 F1 C5 48 37 C6 60 99 74 04 4D 0C " +
                    "BE 68 04 E6 83 6F 75 EB FE C4 B5 5D DA CA C6 B7 " +
                    "E7 4E 20 F0 D4 9E DB FB 3B F7 9E CE 17 BD 6C C8 " +
                    "4E 18 5B 6B C5 C0 1F 79 40 84 CA 74 5E AB 92 41 " +
                    "BA 7D 7D 8C 80 58 C7 55 CE 1B AD 04 83 BD B0 B0 " +
                    "48 6E 6F E6 3C CE 75 A7 D6 62 B7 71 25 41 13 65 " +
                    "63 CE A0 72 AE 8D 29 B5 59 25 24 B7 23 DD 0A 85 " +
                    "6F 85 36 92 3B 94 A6 A4 B9 E1 1D 9A CB 86 8E 3D " +
                    "6F 4A 25 AF 15 4E A8 78 2B 5E 2A 91 FD 6E FE 24 " +
                    "0E F9 94 5F A9 7E 8E BE DF E9 43 DA AD 57 DB A8 " +
                    "0E 83 88 B1 FB BB 7E F9 04 C4 89 DE 5D 9A 4F F4 " +
                    "5B CE 60 0A A4 58 ED 7F 4C 9D AF B9 75 C2 30 C0 " +
                    "38 18 0E 83 41 82 1B F7 CD 52 65 95 36 C3 BD 30 " +
                    "5A 12 A3 3B 06 33 20 99 6E 18 04 30 E8 8F A2 B0 " +
                    "C2 FD 93 58 3D AB C0 9F 04 E3 10 E8 40 3A 7D E4 " +
                    "FC F0 04 E2 79 05 4E 66 53 EC 1C 2C CF 70 18 F9 " +
                    "DE 11 A6 D7 3B A0 B8 7C 61 72 00 00 00 FF FF 03 " +
                    "00 50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 21 " +
                    "00 F0 F7 8A BB FD 00 00 00 E2 01 00 00 13 00 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 5B " +
                    "43 6F 6E 74 65 6E 74 5F 54 79 70 65 73 5D 2E 78 " +
                    "6D 6C 50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 " +
                    "21 00 31 DD 5F 61 D2 00 00 00 8F 01 00 00 0B 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 2E 01 00 00 " +
                    "5F 72 65 6C 73 2F 2E 72 65 6C 73 50 4B 01 02 2D " +
                    "00 14 00 06 00 08 00 00 00 21 00 D1 22 B3 C1 96 " +
                    "02 00 00 F9 06 00 00 10 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 29 02 00 00 64 72 73 2F 73 68 61 " +
                    "70 65 78 6D 6C 2E 78 6D 6C 50 4B 01 02 2D 00 14 " +
                    "00 06 00 08 00 00 00 21 00 EA 0B 8D A0 1A 01 00 " +
                    "00 8A 01 00 00 0F 00 00 00 00 00 00 00 00 00 00 " +
                    "00 00 00 ED 04 00 00 64 72 73 2F 64 6F 77 6E 72 " +
                    "65 76 2E 78 6D 6C 50 4B 05 06 00 00 00 00 04 00 " +
                    "04 00 F5 00 00 00 34 06 00 00 00 00 00 00 10 F0 " +
                    "12 00 00 00 00 00 03 00 10 02 07 00 00 00 05 00 " +
                    "D0 03 0F 00 40 00 00 00 11 F0 00 00 00 00 5D 00 " +
                    "1A 00 15 00 12 00 02 00 05 00 11 60 00 00 00 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 3C 00 E7 07 " +
                    "0F 00 04 F0 DF 07 00 00 12 00 0A F0 08 00 00 00 " +
                    "06 04 00 00 00 0A 00 00 83 00 0B F0 50 00 00 00 " +
                    "BF 00 18 00 1F 00 81 01 4F 81 BD 00 BF 01 10 00 " +
                    "10 00 C0 01 38 5D 8A 00 CB 01 38 63 00 00 FF 01 " +
                    "08 00 08 00 80 C3 20 00 00 00 BF 03 00 00 02 00 " +
                    "1F 04 40 04 4F 04 3C 04 3E 04 43 04 33 04 3E 04 " +
                    "3B 04 4C 04 3D 04 38 04 3A 04 20 00 36 00 00 00 " +
                    "23 00 22 F1 4D 07 00 00 FF 01 00 00 40 00 A9 C3 " +
                    "41 07 00 00 50 4B 03 04 14 00 06 00 08 00 00 00 " +
                    "21 00 F0 F7 8A BB FD 00 00 00 E2 01 00 00 13 00 " +
                    "00 00 5B 43 6F 6E 74 65 6E 74 5F 54 79 70 65 73 " +
                    "5D 2E 78 6D 6C 94 91 CD 4A C4 30 10 C7 EF 82 EF " +
                    "10 E6 2A 6D AA 07 11 69 BA 07 AB 47 15 5D 1F 60 " +
                    "48 A6 6D D8 36 09 99 58 77 DF DE 74 3F 2E E2 0A " +
                    "1E 67 E6 FF F1 23 A9 57 DB 69 14 33 45 B6 DE 29 " +
                    "B8 2E 2B 10 E4 B4 37 D6 F5 0A 3E D6 4F C5 1D 08 " +
                    "4E E8 0C 8E DE 91 82 1D 31 AC 9A CB 8B 7A BD 0B " +
                    "C4 22 BB 1D 2B 18 52 0A F7 52 B2 1E 68 42 2E 7D " +
                    "20 97 2F 9D 8F 13 A6 3C C6 5E 06 D4 1B EC 49 DE " +
                    "54 D5 AD D4 DE 25 72 A9 48 4B 06 34 75 4B 1D 7E " +
                    "8E 49 3C 6E F3 FA 40 12 69 64 10 0F 07 E1 D2 A5 " +
                    "00 43 18 AD C6 94 49 E5 EC CC 8F 96 E2 D8 50 66 " +
                    "E7 5E C3 83 0D 7C 95 31 40 FE DA B0 5C CE 17 1C " +
                    "7D 2F F9 69 A2 35 24 5E 31 A6 67 9C 32 86 34 91 " +
                    "25 0F 18 28 6B CA BF 53 16 CC 89 0B DF 75 56 53 " +
                    "D9 46 7E 5F 7C 27 A8 73 E1 C6 7F B9 48 F3 7F B3 " +
                    "DB 6C 7B A3 F9 94 2E F7 3F D4 7C 03 00 00 FF FF " +
                    "03 00 50 4B 03 04 14 00 06 00 08 00 00 00 21 00 " +
                    "31 DD 5F 61 D2 00 00 00 8F 01 00 00 0B 00 00 00 " +
                    "5F 72 65 6C 73 2F 2E 72 65 6C 73 A4 90 C1 6A C3 " +
                    "30 0C 86 EF 83 BD 83 D1 BD 71 DA 43 19 A3 4E 6F " +
                    "85 5E 4B 07 BB 0A 5B 49 4C 63 CB 58 26 6D DF BE " +
                    "A6 30 58 46 6F 3B EA 17 FA 3E F1 EF F6 B7 30 A9 " +
                    "99 B2 78 8E 06 D6 4D 0B 8A A2 65 E7 E3 60 E0 EB " +
                    "7C 58 7D 80 92 82 D1 E1 C4 91 0C DC 49 60 DF BD " +
                    "BF ED 4E 34 61 A9 47 32 FA 24 AA 52 A2 18 18 4B " +
                    "49 9F 5A 8B 1D 29 A0 34 9C 28 D6 4D CF 39 60 A9 " +
                    "63 1E 74 42 7B C1 81 F4 A6 6D B7 3A FF 66 40 B7 " +
                    "60 AA A3 33 90 8F 6E 03 EA 7C 4F D5 FC 87 1D BC " +
                    "CD 2C DC 97 C6 72 D0 DC F7 DE BE A2 6A C7 D7 78 " +
                    "A2 B9 52 30 0F 54 0C B8 2C CF 30 D3 DC D4 E7 40 " +
                    "BF F6 AE FF E9 95 11 13 7D 57 FE 42 FC 4C AB F5 " +
                    "C7 AC 17 35 76 0F 00 00 00 FF FF 03 00 50 4B 03 " +
                    "04 14 00 06 00 08 00 00 00 21 00 9D C5 0E FA 96 " +
                    "02 00 00 FD 06 00 00 10 00 00 00 64 72 73 2F 73 " +
                    "68 61 70 65 78 6D 6C 2E 78 6D 6C AC 55 4B 6E DB " +
                    "30 10 DD 17 E8 1D 08 EE 13 49 8E E5 8F 60 29 68 " +
                    "1D B4 9B A2 31 9C E6 00 AC 44 D9 42 29 52 20 59 " +
                    "5B CE AA 40 B7 05 7A 84 1E A2 9B A2 9F 9C 41 BE " +
                    "51 87 A4 A4 B4 E9 67 11 DB 0B 9B 9E 21 E7 BD 79 " +
                    "33 43 CE CE EB 92 A1 0D 95 AA 10 3C C6 C1 A9 8F " +
                    "11 E5 A9 C8 0A BE 8A F1 F5 AB 67 27 13 8C 94 26 " +
                    "3C 23 4C 70 1A E3 1D 55 F8 3C 79 FC 68 56 67 32 " +
                    "22 3C 5D 0B 89 20 04 57 11 18 62 BC D6 BA 8A 3C " +
                    "4F A5 6B 5A 12 75 2A 2A CA C1 9B 0B 59 12 0D 7F " +
                    "E5 CA CB 24 D9 42 F0 92 79 03 DF 1F 79 AA 92 94 " +
                    "64 6A 4D A9 BE 70 1E 9C D8 D8 7A 2B E6 94 B1 27 " +
                    "16 C2 99 72 29 4A B7 4A 05 4B CE 66 9E E1 60 96 " +
                    "F6 00 2C 2E F3 3C 19 8E 46 E3 41 D8 FB 8C C9 BA " +
                    "A5 D8 26 63 67 36 CB CE 66 FC 41 38 18 FA 7E EF " +
                    "B3 47 6C EC 3B 40 2D 7A 90 64 D4 07 EF 6D 36 CA " +
                    "E0 6C F2 2F E0 A0 3D 73 1F 79 1A 76 27 C0 73 87 " +
                    "DB A1 A9 0A 95 24 95 22 C6 18 69 5A 6B 56 F0 37 " +
                    "B0 76 B0 7C 73 55 2D 64 4B E1 E5 66 21 51 91 C5 " +
                    "78 8C 11 27 25 14 AA F9 B4 7F B7 FF D8 7C 6F 6E " +
                    "F7 EF 9B CF CD 6D F3 6D FF A1 F9 D1 7C 69 BE A2 " +
                    "11 F6 FA 63 26 06 FC B3 E9 FE 1A 51 D9 D8 24 AA " +
                    "73 59 B6 05 26 0F 28 6F 49 0A 0E 7C 49 24 F2 1C " +
                    "D5 31 1E 0C A6 21 E4 8C D1 0E 9A 6D 38 09 A7 BE " +
                    "6F C8 90 08 B2 43 29 6C E8 8C 28 35 3B C2 71 30 " +
                    "82 DD 86 A0 A3 62 B6 56 52 E9 E7 54 1C 4C 0B 99 " +
                    "40 31 96 34 D5 96 22 D9 BC 50 DA 41 75 10 AD 30 " +
                    "4E 0C D3 6D 4A EF 18 35 24 18 5F 52 C8 C8 76 FE " +
                    "83 85 81 8A 19 4D 2C BA 1D 19 3A 67 12 6D 08 8B " +
                    "31 49 53 CA 75 E0 5C 6B 92 51 67 0E 7D F8 B4 7A " +
                    "F4 27 AC 3A 96 90 61 96 17 8C 1D 8D 5B 4B C0 8C " +
                    "F3 9F DC 9C 56 2D 9E 2D 62 9E 83 98 47 03 F7 FF " +
                    "27 8C 03 A7 1D A2 CD 5C F0 E3 81 97 05 17 F2 6F " +
                    "04 18 54 A5 CD DC E1 75 4D E2 5A C3 74 89 AE 9F " +
                    "8A 6C 67 28 BD 86 5F 98 CC 43 FB 04 2E 68 7D 09 " +
                    "5F 39 13 DB 18 A7 AC A8 30 82 9B F7 E6 BE 4D 6A " +
                    "36 17 D0 3D 70 8B BB BB 39 C6 DA CD 17 53 FA CA " +
                    "10 3C 94 0A 64 0E 13 78 68 14 1B 04 74 21 6C 05 " +
                    "AF 0E 73 14 29 CF 16 44 92 25 D8 19 31 CF 8F 7C " +
                    "7B B2 BC 86 E7 E7 06 6E 82 A0 6F FB AA D5 BB 13 " +
                    "D9 5E 5D 0A AC F6 35 60 05 8C CD 05 D1 C4 94 C8 " +
                    "D6 E2 F7 77 C4 DA 9C 36 C9 4F 00 00 00 FF FF 03 " +
                    "00 50 4B 03 04 14 00 06 00 08 00 00 00 21 00 A3 " +
                    "9B 83 A2 1C 01 00 00 8E 01 00 00 0F 00 00 00 64 " +
                    "72 73 2F 64 6F 77 6E 72 65 76 2E 78 6D 6C 4C 90 " +
                    "CD 4E C3 30 10 84 EF 48 BC 83 B5 48 5C 10 75 9A " +
                    "34 69 08 75 AA 82 84 DA 0B 88 B6 E1 C0 CD 24 CE " +
                    "8F 88 ED CA 36 4D FA F6 6C 5B 50 73 B2 C7 3B DF " +
                    "EE 8E 67 F3 5E B6 64 2F 8C 6D B4 62 30 1E 79 40 " +
                    "84 CA 75 D1 A8 8A 41 B6 7D B9 8F 81 58 C7 55 C1 " +
                    "5B AD 04 83 83 B0 30 4F AF AF 66 3C 29 74 A7 D6 " +
                    "62 BF 71 15 C1 26 CA 26 9C 41 ED DC 2E A1 D4 E6 " +
                    "B5 90 DC 8E F4 4E 28 AC 95 DA 48 EE 50 9A 8A 16 " +
                    "86 77 D8 5C B6 D4 F7 BC 88 4A DE 28 9C 50 F3 9D " +
                    "78 AE 45 FE BD F9 91 38 E4 5D 7E 64 FA 29 FE 7C " +
                    "A5 77 59 B7 5E 6E E3 26 0C 62 C6 6E 6F FA C5 23 " +
                    "10 27 7A 77 31 FF D1 AB 82 C1 14 48 B9 3C 7C 99 " +
                    "A6 58 73 EB 84 61 80 71 30 1C 06 83 14 37 EE DB " +
                    "85 CA 6B 6D 8E F7 D2 68 49 8C EE 4E 54 AE 5B 06 " +
                    "01 1C F5 5B 59 5A E1 90 08 FD 89 87 38 96 FE 9F " +
                    "26 51 34 F5 43 A0 47 DC E9 33 3C 8E 4E 16 06 78 " +
                    "0E E8 87 10 8D 43 76 EC 07 F1 99 A5 C3 3D 50 5C " +
                    "BE 31 FD 05 00 00 FF FF 03 00 50 4B 01 02 2D 00 " +
                    "14 00 06 00 08 00 00 00 21 00 F0 F7 8A BB FD 00 " +
                    "00 00 E2 01 00 00 13 00 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 00 00 00 5B 43 6F 6E 74 65 6E 74 " +
                    "5F 54 79 70 65 73 5D 2E 78 6D 6C 50 4B 01 02 2D " +
                    "00 14 00 06 00 08 00 00 00 21 00 31 DD 5F 61 D2 " +
                    "00 00 00 8F 01 00 00 0B 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 2E 01 00 00 5F 72 65 6C 73 2F 2E " +
                    "72 65 6C 73 50 4B 01 02 2D 00 14 00 06 00 08 00 " +
                    "00 00 21 00 9D C5 0E FA 96 02 00 00 FD 06 00 00 " +
                    "10 00 00 00 00 00 00 00 00 00 00 00 00 00 29 02 " +
                    "00 00 64 72 73 2F 73 68 61 70 65 78 6D 6C 2E 78 " +
                    "6D 6C 50 4B 01 02 2D 00 14 00 06 00 08 00 00 00 " +
                    "21 00 A3 9B 83 A2 1C 01 00 00 8E 01 00 00 0F 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 ED 04 00 00 " +
                    "64 72 73 2F 64 6F 77 6E 72 65 76 2E 78 6D 6C 50 " +
                    "4B 05 06 00 00 00 00 04 00 04 00 F5 00 00 00 36 " +
                    "06 00 00 00 00 00 00 10 F0 12 00 00 00 00 00 03 " +
                    "00 10 03 07 00 CD 00 06 00 D0 00 10 00 0D 00 00 " +
                    "00 11 F0 00 00 00 00 5D 00 1A 00 15 00 12 00 02 " +
                    "00 06 00 11 60 00 00 00 00 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 3C 00 E8 07 0F 00 04 F0 E0 07 00 " +
                    "00 12 00 0A F0 08 00 00 00 07 04 00 00 00 0A 00 " +
                    "00 83 00 0B F0 50 00 00 00 BF 00 18 00 1F 00 81 " +
                    "01 4F 81 BD 00 BF 01 10 00 10 00 C0 01 38 5D 8A " +
                    "00 CB 01 38 63 00 00 FF 01 08 00 08 00 80 C3 20 " +
                    "00 00 00 BF 03 00 00 02 00 1F 04 40 04 4F 04 3C " +
                    "04 3E 04 43 04 33 04 3E 04 3B 04 4C 04 3D 04 38 " +
                    "04 3A 04 20 00 37 00 00 00 23 00 22 F1 4E 07 00 " +
                    "00 FF 01 00 00 40 00 A9 C3 42 07 00 00 50 4B 03 " +
                    "04 14 00 06 00 08 00 00 00 21 00 F0 F7 8A BB FD " +
                    "00 00 00 E2 01 00 00 13 00 00 00 5B 43 6F 6E 74 " +
                    "65 6E 74 5F 54 79 70 65 73 5D 2E 78 6D 6C 94 91 " +
                    "CD 4A C4 30 10 C7 EF 82 EF 10 E6 2A 6D AA 07 11 " +
                    "69 BA 07 AB 47 15 5D 1F 60 48 A6 6D D8 36 09 99 " +
                    "58 77 DF DE 74 3F 2E E2 0A 1E 67 E6 FF F1 23 A9 " +
                    "57 DB 69 14 33 45 B6 DE 29 B8 2E 2B 10 E4 B4 37 " +
                    "D6 F5 0A 3E D6 4F C5 1D 08 4E E8 0C 8E DE 91 82 " +
                    "1D 31 AC 9A CB 8B 7A BD 0B C4 22 BB 1D 2B 18 52 " +
                    "0A F7 52 B2 1E 68 42 2E 7D 20 97 2F 9D 8F 13 A6 " +
                    "3C C6 5E 06 D4 1B EC 49 DE 54 D5 AD D4 DE 25 72 " +
                    "A9 48 4B 06 34 75 4B 1D 7E 8E 49 3C 6E F3 FA 40 " +
                    "12 69 64 10 0F 07 E1 D2 A5 00 43 18 AD C6 94 49 " +
                    "E5 EC CC 8F 96 E2 D8 50 66 E7 5E C3 83 0D 7C 95 " +
                    "31 40 FE DA B0 5C CE 17 1C 7D 2F F9 69 A2 35 24 " +
                    "5E 31 A6 67 9C 32 86 34 91 25 0F 18 28 6B CA BF " +
                    "53 16 CC 89 0B DF 75 56 53 D9 46 7E 5F 7C 27 A8 " +
                    "73 E1 C6 7F B9 48 F3 7F B3 DB 6C 7B A3 F9 94 2E " +
                    "F7 3F D4 7C 03 00 00 FF FF 03 00 50 4B 03 04 14 " +
                    "00 06 00 08 00 00 00 21 00 31 DD 5F 61 D2 00 00 " +
                    "00 8F 01 00 00 0B 00 00 00 5F 72 65 6C 73 2F 2E " +
                    "72 65 6C 73 A4 90 C1 6A C3 30 0C 86 EF 83 BD 83 " +
                    "D1 BD 71 DA 43 19 A3 4E 6F 85 5E 4B 07 BB 0A 5B " +
                    "49 4C 63 CB 58 26 6D DF BE A6 30 58 46 6F 3B EA " +
                    "17 FA 3E F1 EF F6 B7 30 A9 99 B2 78 8E 06 D6 4D " +
                    "0B 8A A2 65 E7 E3 60 E0 EB 7C 58 7D 80 92 82 D1 " +
                    "E1 C4 91 0C DC 49 60 DF BD BF ED 4E 34 61 A9 47 " +
                    "32 FA 24 AA 52 A2 18 18 4B 49 9F 5A 8B 1D 29 A0 " +
                    "34 9C 28 D6 4D CF 39 60 A9 63 1E 74 42 7B C1 81 " +
                    "F4 A6 6D B7 3A FF 66 40 B7 60 AA A3 33 90 8F 6E " +
                    "03 EA 7C 4F D5 FC 87 1D BC CD 2C DC 97 C6 72 D0 " +
                    "DC F7 DE BE A2 6A C7 D7 78 A2 B9 52 30 0F 54 0C " +
                    "B8 2C CF 30 D3 DC D4 E7 40 BF F6 AE FF E9 95 11 " +
                    "13 7D 57 FE 42 FC 4C AB F5 C7 AC 17 35 76 0F 00 " +
                    "00 00 FF FF 03 00 50 4B 03 04 14 00 06 00 08 00 " +
                    "00 00 21 00 33 C2 C5 D0 98 02 00 00 FD 06 00 00 " +
                    "10 00 00 00 64 72 73 2F 73 68 61 70 65 78 6D 6C " +
                    "2E 78 6D 6C AC 55 CD 8E D3 30 10 BE 23 F1 0E 96 " +
                    "EF BB 49 BA 6D DA 46 4D 56 D0 15 5C 10 5B 75 D9 " +
                    "07 30 89 D3 46 38 76 64 9B 36 DD 13 12 57 24 1E " +
                    "81 87 E0 82 F8 D9 67 48 DF 88 B1 9D 64 81 05 0E " +
                    "DB F6 D0 BA 33 F6 7C DF 7C 33 63 CF CE EB 92 A1 " +
                    "0D 95 AA 10 3C C6 C1 A9 8F 11 E5 A9 C8 0A BE 8A " +
                    "F1 F5 AB 67 27 13 8C 94 26 3C 23 4C 70 1A E3 1D " +
                    "55 F8 3C 79 FC 68 56 67 32 22 3C 5D 0B 89 20 04 " +
                    "57 11 18 62 BC D6 BA 8A 3C 4F A5 6B 5A 12 75 2A " +
                    "2A CA C1 9B 0B 59 12 0D 7F E5 CA CB 24 D9 42 F0 " +
                    "92 79 03 DF 0F 3D 55 49 4A 32 B5 A6 54 5F 38 0F " +
                    "4E 6C 6C BD 15 73 CA D8 13 0B E1 4C B9 14 A5 5B " +
                    "A5 82 25 C3 99 67 38 98 A5 3D 00 8B CB 3C 4F A6 " +
                    "A3 C1 A8 F7 18 83 75 4A B1 4D 26 CE 6C 96 9D CD " +
                    "F8 83 60 78 E6 FB BD CF 1E B1 91 EF E0 B4 E8 21 " +
                    "92 B0 0F DE DB CC 91 C1 38 1C FC 0B 38 68 CF DC " +
                    "43 0E 83 69 77 06 7C 77 C8 1D 9E AA 50 49 52 29 " +
                    "62 8C 91 A6 B5 66 05 7F 03 6B 07 CC 37 57 D5 42 " +
                    "B6 24 5E 6E 16 12 15 59 8C A1 5C 9C 94 50 A8 E6 " +
                    "D3 FE DD FE 63 F3 BD B9 DD BF 6F 3E 37 B7 CD B7 " +
                    "FD 87 E6 47 F3 A5 F9 8A C6 D8 EB 8F 99 18 F0 CF " +
                    "26 FC 6B 44 65 63 93 A8 CE 65 D9 16 98 3C A0 BC " +
                    "25 29 38 F0 25 91 C8 73 54 C7 78 30 1C 8E 21 67 " +
                    "8C 76 D0 6C E1 D9 04 94 37 64 48 04 D9 A1 14 36 " +
                    "04 C3 C9 68 0A 46 94 9A 1D A3 71 10 C2 6E 43 D0 " +
                    "51 31 5B 2B A9 F4 73 2A 0E A6 85 4C A0 18 4B 9A " +
                    "6A 4B 91 6C 5E 28 ED A0 3A 88 56 18 27 86 E9 36 " +
                    "A5 77 8C 1A 12 8C 2F 29 64 64 3B FF C1 C2 40 C5 " +
                    "8C 26 16 DD 8E 0C 9D 33 89 36 84 C5 98 A4 29 E5 " +
                    "3A 70 AE 35 C9 A8 33 8F 7C F8 B4 7A F4 27 AC 3A " +
                    "96 90 61 96 17 8C 1D 8D 5B 4B C0 8C F3 7D 6E 4E " +
                    "AB 16 CF 16 31 CF 41 CC A3 81 FB FF 13 C6 81 D3 " +
                    "0E D1 66 2E F8 F1 C0 CB 82 0B F9 37 02 0C AA D2 " +
                    "66 EE F0 BA 26 71 AD 61 BA 44 D7 4F 45 B6 33 94 " +
                    "5E C3 2F 4C E6 A1 7D 02 17 B4 BE 84 AF 9C 89 6D " +
                    "8C 53 56 54 18 C1 CD 7B F3 A7 4D 6A 36 17 D0 3D " +
                    "30 3F EE 6E 8E B1 76 F3 C5 94 BE 32 04 0F A5 02 " +
                    "99 C3 04 1E 1A C5 06 01 5D 08 5B C1 AB C3 1C 45 " +
                    "CA B3 05 91 64 09 76 46 CC F3 23 DF 9E 2C AF E1 " +
                    "F9 B9 81 9B 20 E8 DB BE 6A F5 EE 44 B6 57 97 02 " +
                    "AB 7D 0D 58 01 63 73 41 34 31 25 B2 B5 F8 FD 1D " +
                    "B1 36 A7 4D F2 13 00 00 FF FF 03 00 50 4B 03 04 " +
                    "14 00 06 00 08 00 00 00 21 00 6F E2 C1 89 1B 01 " +
                    "00 00 8E 01 00 00 0F 00 00 00 64 72 73 2F 64 6F " +
                    "77 6E 72 65 76 2E 78 6D 6C 4C 90 CD 4E C3 30 10 " +
                    "84 EF 48 BC 43 B4 48 5C 10 75 92 36 21 0D 75 AA " +
                    "82 84 DA 0B 88 B4 E1 C0 CD 24 CE 8F 88 ED C8 36 " +
                    "4D FA F6 6C 5A 50 7B B2 C7 BB DF EE 8C 17 CB 41 " +
                    "B4 CE 9E 6B D3 28 49 C1 9B B8 E0 70 99 AB A2 91 " +
                    "15 85 6C F7 72 1F 81 63 2C 93 05 6B 95 E4 14 0E " +
                    "DC C0 32 B9 BE 5A B0 B8 50 BD 4C F9 7E 6B 2B 07 " +
                    "87 48 13 33 0A B5 B5 5D 4C 88 C9 6B 2E 98 99 A8 " +
                    "8E 4B AC 95 4A 0B 66 51 EA 8A 14 9A F5 38 5C B4 " +
                    "C4 77 DD 90 08 D6 48 DC 50 B3 8E 3F D7 3C FF DE " +
                    "FE 08 5C F2 2E 3E 32 F5 14 7D BE 92 BB AC 4F D7 " +
                    "BB A8 09 A6 11 A5 B7 37 C3 EA 11 1C CB 07 7B 6E " +
                    "FE A3 37 05 05 F4 5A AE 0F 5F BA 29 52 66 2C D7 " +
                    "14 30 0E 86 C3 60 90 A0 E3 A1 5D C9 BC 56 7A BC " +
                    "97 5A 09 47 AB FE 48 E5 AA A5 30 83 51 BF 95 A5 " +
                    "E1 16 09 6F 36 75 11 C7 D2 FF D3 3C F0 03 20 23 " +
                    "6C D5 09 F5 C2 63 03 05 3C 2F D9 D0 9B 63 EB 25 " +
                    "EB 3F 84 FE 89 26 97 3E 50 9C BF 31 F9 05 00 00 " +
                    "FF FF 03 00 50 4B 01 02 2D 00 14 00 06 00 08 00 " +
                    "00 00 21 00 F0 F7 8A BB FD 00 00 00 E2 01 00 00 " +
                    "13 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                    "00 00 5B 43 6F 6E 74 65 6E 74 5F 54 79 70 65 73 " +
                    "5D 2E 78 6D 6C 50 4B 01 02 2D 00 14 00 06 00 08 " +
                    "00 00 00 21 00 31 DD 5F 61 D2 00 00 00 8F 01 00 " +
                    "00 0B 00 00 00 00 00 00 00 00 00 00 00 00 00 2E " +
                    "01 00 00 5F 72 65 6C 73 2F 2E 72 65 6C 73 50 4B " +
                    "01 02 2D 00 14 00 06 00 08 00 00 00 21 00 33 C2 " +
                    "C5 D0 98 02 00 00 FD 06 00 00 10 00 00 00 00 00 " +
                    "00 00 00 00 00 00 00 00 29 02 00 00 64 72 73 2F " +
                    "73 68 61 70 65 78 6D 6C 2E 78 6D 6C 50 4B 01 02 " +
                    "2D 00 14 00 06 00 08 00 00 00 21 00 6F E2 C1 89 " +
                    "1B 01 00 00 8E 01 00 00 0F 00 00 00 00 00 00 00 " +
                    "00 00 00 00 00 00 EF 04 00 00 64 72 73 2F 64 6F " +
                    "77 6E 72 65 76 2E 78 6D 6C 50 4B 05 06 00 00 00 " +
                    "00 04 00 04 00 F5 00 00 00 37 06 00 00 00 00 00 " +
                    "00 10 F0 12 00 00 00 00 00 04 00 10 00 08 00 9A " +
                    "00 06 00 D0 01 10 00 DA 00 00 00 11 F0 00 00 00 " +
                    "00 5D 00 1A 00 15 00 12 00 02 00 07 00 11 60 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00";


            byte[] dgBytes = HexRead.ReadFromString(data);
            IList<NPOI.HSSF.Record.Record> dgRecords = RecordFactory.CreateRecords(new MemoryStream(dgBytes));
            Assert.AreEqual(14, dgRecords.Count);

            short[] expectedSids = {
                    DrawingRecord.sid,
                    ObjRecord.sid,
                    DrawingRecord.sid,
                    ObjRecord.sid,
                    DrawingRecord.sid,
                    ObjRecord.sid,
                    DrawingRecord.sid,
                    ObjRecord.sid,
                    ContinueRecord.sid,
                    ObjRecord.sid,
                    ContinueRecord.sid,
                    ObjRecord.sid,
                    ContinueRecord.sid,
                    ObjRecord.sid
            };

            for (int i = 0; i < expectedSids.Length; i++)
            {
                Assert.AreEqual(expectedSids[i], dgRecords[(i)].Sid, "unexpected record.sid and index[" + i + "]");
            }
            DrawingManager2 drawingManager = new DrawingManager2(new EscherDggRecord());

            // create a dummy sheet consisting of our Test data
            InternalSheet sheet = InternalSheet.CreateSheet();
            List<RecordBase> records = sheet.Records;
            records.Clear();
            records.AddRange(dgRecords);
            records.Add(EOFRecord.instance);

            sheet.AggregateDrawingRecords(drawingManager, false);
            Assert.AreEqual(2, records.Count, "drawing was not fully aggregated");
            Assert.IsTrue(records[(0)] is EscherAggregate , "expected EscherAggregate");
            Assert.IsTrue(records[(1)] is EOFRecord , "expected EOFRecord");

            EscherAggregate agg = (EscherAggregate)records[(0)];

            byte[] dgBytesAfterSave = agg.Serialize();
            Assert.AreEqual(dgBytes.Length, dgBytesAfterSave.Length, "different size of Drawing data before and After save");
            Assert.IsTrue(Arrays.Equals(dgBytes, dgBytesAfterSave), "drawing data brefpore and After save is different");
        }


    }
}
