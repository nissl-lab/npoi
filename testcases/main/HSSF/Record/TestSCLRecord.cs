
/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */



namespace TestCases.HSSF.Record
{
    using System;
    using NPOI.HSSF.Record;

    using NUnit.Framework;

    /**
     * Tests the serialization and deserialization of the SCLRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Andrew C. Oliver (acoliver at apache.org)
     */
    [TestFixture]
    public class TestSCLRecord
    {
        byte[] data = new byte[] {
      (byte)0x3,(byte)0x0,(byte)0x4,(byte)0x0
    };

        [Test]
        public void TestLoad()
        {
            SCLRecord record = new SCLRecord(TestcaseRecordInputStream.Create((short)0xa0, data));
            Assert.AreEqual(3, record.Numerator);
            Assert.AreEqual(4, record.Denominator);


            Assert.AreEqual(8, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            SCLRecord record = new SCLRecord();
            record.Numerator = ((short)3);
            record.Denominator = ((short)4);


            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}