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
    using System;


    using NUnit.Framework;
    using TestCases.HSSF.Record;
    using NPOI.Util;
    using NPOI.HSSF.Record;



    /**
     * Tests the serialization and deserialization of the FtCblsSubRecord
     * class works correctly.
     *
     * @author Yegor Kozlov
     */
    [TestFixture]
    public class TestFtCblsSubRecord
    {
        private byte[] data = new byte[] {
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x64, 0x00,
        0x01, 0x00, 0x0A, 0x00, 0x00, 0x00, 0x10, 0x00, 0x01, 0x00
    };
        [Test]
        public void TestRead()
        {

            FtCblsSubRecord record = new FtCblsSubRecord(TestcaseRecordInputStream.Create(FtCblsSubRecord.sid, data), data.Length);

            Assert.AreEqual(FtCblsSubRecord.sid, record.Sid);
            Assert.AreEqual(data.Length, record.DataSize);
        }
        [Test]
        public void TestWrite()
        {
            FtCblsSubRecord record = new FtCblsSubRecord();
            Assert.AreEqual(FtCblsSubRecord.sid, record.Sid);
            Assert.AreEqual(data.Length, record.DataSize);

            byte[] ser = record.Serialize();
            Assert.AreEqual(ser.Length - 4, data.Length);

        }
        [Test]
        public void TestClone()
        {
            FtCblsSubRecord record = new FtCblsSubRecord();
            byte[] src = record.Serialize();

            FtCblsSubRecord Cloned = (FtCblsSubRecord)record.Clone();
            byte[] cln = Cloned.Serialize();

            Assert.AreEqual(record.DataSize, Cloned.DataSize);
            Assert.IsTrue(Arrays.Equals(src, cln));
        }
    }
}