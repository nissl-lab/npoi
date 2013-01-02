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

namespace TestCases.HSSF.Record
{

    using NUnit.Framework;
    using NPOI.HSSF.Record.AutoFilter;
    using NPOI.Util;

    /**
     * Tests the AutoFilterInfoRecord class.
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestAutoFilterInfoRecord
    {
        private byte[] data = new byte[] {
        0x05, 0x00
    };
        [Test]
        public void TestRead()
        {

            AutoFilterInfoRecord record = new AutoFilterInfoRecord(TestcaseRecordInputStream.Create(AutoFilterInfoRecord.sid, data));

            Assert.AreEqual(AutoFilterInfoRecord.sid, record.Sid);
            Assert.AreEqual(data.Length, record.RecordSize - 4);
            Assert.AreEqual(5, record.NumEntries);
            record.NumEntries = (/*setter*/(short)3);
            Assert.AreEqual(3, record.NumEntries);
        }
        [Test]
        public void TestWrite()
        {
            AutoFilterInfoRecord record = new AutoFilterInfoRecord();
            record.NumEntries = (/*setter*/(short)3);

            byte[] ser = record.Serialize();
            Assert.AreEqual(ser.Length - 4, data.Length);
            record = new AutoFilterInfoRecord(TestcaseRecordInputStream.Create(ser));
            Assert.AreEqual(3, record.NumEntries);
        }
        [Test]
        public void TestClone()
        {
            AutoFilterInfoRecord record = new AutoFilterInfoRecord();
            record.NumEntries = (/*setter*/(short)3);
            byte[] src = record.Serialize();

            AutoFilterInfoRecord Cloned = (AutoFilterInfoRecord)record.Clone();
            Assert.AreEqual(3, record.NumEntries);
            byte[] cln = Cloned.Serialize();

            Assert.AreEqual(record.RecordSize, Cloned.RecordSize);
            Assert.IsTrue(Arrays.Equals(src, cln));
        }
    }
}