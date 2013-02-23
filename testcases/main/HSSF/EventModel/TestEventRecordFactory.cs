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

namespace TestCases.HSSF.EventModel
{
    using System;
    using System.IO;
    using System.Collections;

    using NPOI.HSSF;
    using NPOI.HSSF.EventModel;
    using TestCases.HSSF.Record;
    using NPOI.HSSF.Record;
    using NPOI.POIFS.FileSystem;

    using NUnit.Framework;


    /**
     * enclosing_type describe the purpose here
     * 
     * @author Andrew C. Oliver acoliver@apache.org
     * @author Csaba Nagy (ncsaba at yahoo dot com)
     */
    [TestFixture]
    public class TestEventRecordFactory
    {


        class ERFListener1 : IERFListener
        {
            bool[] wascalled;
            public ERFListener1(ref bool[] wascalled)
            {
                this.wascalled = wascalled;
            }

            public bool ProcessRecord(Record rec)
            {
                wascalled[0] = true;
                Assert.IsTrue(
                           (rec.Sid == BOFRecord.sid),
                           "must be BOFRecord got SID=" + rec.Sid);
                return true;
            }
        }

        /**
         * Tests that the records can be Processed and properly return 
         * values.
         */
        [Test]
        public void TestProcessRecords()
        {
            bool[] wascalled = { false, }; // hack to pass boolean by ref into inner class

            IERFListener listener = new ERFListener1(ref wascalled);
            ArrayList param = new ArrayList();
            param.Add(BOFRecord.sid);
            EventRecordFactory factory = new EventRecordFactory(listener, param);

            BOFRecord bof = new BOFRecord();
            bof.Build = ((short)0);
            bof.BuildYear = ((short)1999);
            bof.RequiredVersion = (123);
            bof.Type = (BOFRecordType.Workbook);
            bof.Version = ((short)0x06);
            bof.HistoryBitMask = (BOFRecord.HISTORY_MASK);

            EOFRecord eof = EOFRecord.instance;
            byte[] bytes = new byte[bof.RecordSize + eof.RecordSize];
            int offset = 0;
            offset = bof.Serialize(offset, bytes);
            offset = eof.Serialize(offset, bytes);

            factory.ProcessRecords(new MemoryStream(bytes));
            Assert.IsTrue(wascalled[0], "The record listener must be called");
        }

        /**
         * Tests that the create record function returns a properly 
         * constructed record in the simple case.
         */
        [Test]
        public void TestCreateRecord()
        {
            BOFRecord bof = new BOFRecord();
            bof.Build = ((short)0);
            bof.BuildYear = ((short)1999);
            bof.RequiredVersion = (123);
            bof.Type = (BOFRecordType.Workbook);
            bof.Version = ((short)0x06);
            bof.HistoryBitMask = (BOFRecord.HISTORY_MASK);

            byte[] bytes = bof.Serialize();

            Record[] records = RecordFactory.CreateRecord(TestcaseRecordInputStream.Create(bytes));

            Assert.IsTrue(records.Length == 1, "record.Length must be 1, was =" + records.Length);
            Assert.IsTrue(CompareRec(bof, records[0]), "record is the same");

        }

        /**
         * Compare the Serialized bytes of two records are equal
         * @param first the first record to Compare
         * @param second the second record to Compare
         * @return boolean whether or not the record where equal
         */
        private static bool CompareRec(Record first, Record second)
        {
            byte[] rec1 = first.Serialize();
            byte[] rec2 = second.Serialize();

            if (rec1.Length != rec2.Length)
            {
                return false;
            }
            for (int k = 0; k < rec1.Length; k++)
            {
                if (rec1[k] != rec2[k])
                {
                    return false;
                }
            }
            return true;
        }


        /**
         * Tests that the create record function returns a properly 
         * constructed record in the case of a continued record.
         * TODO - need a real world example to put in a unit Test
         */
        public void TestCreateContinuedRecord()
        {
            //  fail("not implemented");
        }



        class ERFListener2 : IERFListener
        {
            private String[] expectedRecordTypes = {
              typeof(UnknownRecord).Name,
              typeof(ContinueRecord).Name,
              typeof(ContinueRecord).Name
          };
            byte[] data;
            int[] recCnt;
            int[] offset;
            public ERFListener2(ref byte[] data, ref int[] recCnt, ref int[] offset)
            {
                this.data = data;
                this.recCnt = recCnt;
                this.offset = offset;
            }
            public bool ProcessRecord(Record rec)
            {
                // System.out.println(rec.toString());
                Assert.AreEqual(
                  expectedRecordTypes[recCnt[0]],
                  rec.GetType().Name,
                  "Record type"
                );
                CompareData(rec, "Record " + recCnt + ": ");
                recCnt[0]++;
                return true;
            }
            private void CompareData(Record record, String message)
            {
                byte[] recData = record.Serialize();
                for (int i = 0; i < recData.Length; i++)
                {
                    Assert.AreEqual(recData[i], data[offset[0]++], message + " data byte " + i);
                }
            }
        }
        /**
         * Test NAME:  Test Creating ContinueRecords After Unknown Records From An InputStream <P>
         * OBJECTIVE:  Test that the RecordFactory given an InputStream
         *             constructs the expected records.<P>
         * SUCCESS:    Record factory creates the expected records.<P>
         * FAILURE:    The wrong records are created or contain the wrong values <P>
         *
         */
        [Test]
        public void TestContinuedUnknownRecord()
        {
            byte[] data = {
            0, unchecked((byte)-1), 0, 0, // an unknown record with 0 Length
            0x3C , 0, 3, 0, 1, 2, 3, // a continuation record with 3 bytes of data
            0x3C , 0, 1, 0, 4 // one more continuation record with 1 byte of data
        };

            int[] recCnt =  {0} ;
            int[] offset =  {0} ;
            IERFListener listener = new ERFListener2(ref data, ref recCnt,ref offset);
            ArrayList sids = new ArrayList(2);
            sids.Add((short)-256);
            sids.Add((short)0x3C);

            EventRecordFactory factory = new EventRecordFactory(listener, sids);

            factory.ProcessRecords(new MemoryStream(data));
            Assert.AreEqual(3, recCnt[0], "nr. of Processed records");
            Assert.AreEqual(data.Length, offset[0], "nr. of Processed bytes");
        }
    }
}