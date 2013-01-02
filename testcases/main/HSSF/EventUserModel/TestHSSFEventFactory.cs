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

    using NUnit.Framework;

    [TestFixture]
    public class TestHSSFEventFactory
    {

        private static Stream OpenSample(String sampleFileName)
        {
            return HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);
        }
        [Test]
        public void TestWithMissingRecords()
        {

            HSSFRequest req = new HSSFRequest();
            MockHSSFListener mockListen = new MockHSSFListener();
            req.AddListenerForAllRecords(mockListen);

            POIFSFileSystem fs = new POIFSFileSystem(OpenSample("SimpleWithSkip.xls"));
            HSSFEventFactory factory = new HSSFEventFactory();
            factory.ProcessWorkbookEvents(req, fs);

            Record[] recs = mockListen.GetRecords();
            // Check we got the records
            Assert.IsTrue(recs.Length > 100);

            // Check that the last few records are as we expect
            // (Makes sure we don't accidently skip the end ones)
            int numRec = recs.Length;
            Assert.AreEqual(typeof(WindowTwoRecord), recs[numRec - 3].GetType());
            Assert.AreEqual(typeof(SelectionRecord), recs[numRec - 2].GetType());
            Assert.AreEqual(typeof(EOFRecord), recs[numRec - 1].GetType());
        }

        public void TestWithCrazyContinueRecords()
        {
            // Some files have crazy ordering of their continue records
            // Check that we don't break on them (bug #42844)

            HSSFRequest req = new HSSFRequest();
            MockHSSFListener mockListen = new MockHSSFListener();
            req.AddListenerForAllRecords(mockListen);

            POIFSFileSystem fs = new POIFSFileSystem(OpenSample("ContinueRecordProblem.xls"));
            HSSFEventFactory factory = new HSSFEventFactory();
            factory.ProcessWorkbookEvents(req, fs);

            Record[] recs = mockListen.GetRecords();
            // Check we got the records
            Assert.IsTrue(recs.Length > 100);

            // And none of them are continue ones
            for (int i = 0; i < recs.Length; i++)
            {
                Assert.IsFalse(recs[i] is ContinueRecord);
            }

            // Check that the last few records are as we expect
            // (Makes sure we don't accidently skip the end ones)
            int numRec = recs.Length;
            Assert.AreEqual(typeof(DVALRecord), recs[numRec - 3].GetType());
            Assert.AreEqual(typeof(DVRecord), recs[numRec - 2].GetType());
            Assert.AreEqual(typeof(EOFRecord), recs[numRec - 1].GetType());
        }

        /**
         * Unknown records can be continued.
         * Check that HSSFEventFactory doesn't break on them.
         * (the Test file was provided in a reopen of bug #42844)
         */
        [Test]
        public void TestUnknownContinueRecords()
        {

            HSSFRequest req = new HSSFRequest();
            MockHSSFListener mockListen = new MockHSSFListener();
            req.AddListenerForAllRecords(mockListen);

            POIFSFileSystem fs = new POIFSFileSystem(OpenSample("42844.xls"));
            HSSFEventFactory factory = new HSSFEventFactory();
            factory.ProcessWorkbookEvents(req, fs);

            Assert.IsTrue(true, "no errors while Processing the file");
        }

        private class MockHSSFListener : IHSSFListener
        {
            private ArrayList records = new ArrayList();

            public MockHSSFListener() { }
            public Record[] GetRecords()
            {
                Record[] result = (Record[])records.ToArray(typeof(Record));
                return result;
            }

            public void ProcessRecord(Record record)
            {
                records.Add(record);
            }
        }
    }
}
