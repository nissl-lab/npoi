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

namespace TestCases.HSSF.EventModel
{
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.EventUserModel;
    using NPOI.HSSF.Record;
    using NPOI.POIFS.FileSystem;
    using NPOI.Util;
    using TestCases.HSSF;

    /**
     * Tests for {@link AbortableHSSFListener}
     */
    [TestFixture]
    public class TestAbortableListener
    {

        private POIFSFileSystem openSample()
        {
            MemoryStream is1 = new MemoryStream(HSSFITestDataProvider.Instance
                    .GetTestDataFileContent("SimpleWithColours.xls"));
            try
            {
                return new POIFSFileSystem(is1);
            }
            catch (IOException e)
            {
                throw new RuntimeException(e);
            }
        }

        [Test]
        public void TestAbortingBasics()
        {
            AbortableCountingListener l = new AbortableCountingListener(1000);

            HSSFRequest req = new HSSFRequest();
            req.AddListenerForAllRecords(l);

            HSSFEventFactory f = new HSSFEventFactory();

            Assert.AreEqual(0, l.countSeen);
            Assert.AreEqual(null, l.lastRecordSeen);

            POIFSFileSystem fs = openSample();
            short res = f.AbortableProcessWorkbookEvents(req, fs);

            Assert.AreEqual(0, res);
            //Assert.AreEqual(175, l.countSeen);
            Assert.AreEqual(176, l.countSeen); //Tony Qu add a sheetext record, so this value should be 176
            Assert.AreEqual(EOFRecord.sid, l.lastRecordSeen.Sid);
        }


        [Test]
        public void TestAbortStops()
        {
            AbortableCountingListener l = new AbortableCountingListener(1);

            HSSFRequest req = new HSSFRequest();
            req.AddListenerForAllRecords(l);

            HSSFEventFactory f = new HSSFEventFactory();

            Assert.AreEqual(0, l.countSeen);
            Assert.AreEqual(null, l.lastRecordSeen);

            POIFSFileSystem fs = openSample();
            short res = f.AbortableProcessWorkbookEvents(req, fs);

            Assert.AreEqual(1234, res);
            Assert.AreEqual(1, l.countSeen);
            Assert.AreEqual(BOFRecord.sid, l.lastRecordSeen.Sid);
        }

        private class AbortableCountingListener : AbortableHSSFListener
        {
            private int abortAfterIndex;
            public int countSeen;
            public Record lastRecordSeen;

            public AbortableCountingListener(int abortAfter)
            {
                abortAfterIndex = abortAfter;
                countSeen = 0;
                lastRecordSeen = null;
            }
            public override short AbortableProcessRecord(Record record)
            {
                countSeen++;
                lastRecordSeen = record;

                if (countSeen == abortAfterIndex)
                {
                    return 1234;
                }
                return 0;
            }
        }
    }
}
