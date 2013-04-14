
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
     * Tests the serialization and deserialization of the CommonObjectDataSubRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestCommonObjectDataSubRecord
    {
        byte[] data = new byte[] {
        (byte)0x12,(byte)0x00,(byte)0x01,(byte)0x00,
        (byte)0x01,(byte)0x00,(byte)0x11,(byte)0x60,
        (byte)0x00,(byte)0x00,(byte)0x00,(byte)0x00,
        (byte)0x00,(byte)0x0D,(byte)0x26,(byte)0x01,
        (byte)0x00,(byte)0x00,
    };

        [Test]
        public void TestLoad()
        {
            CommonObjectDataSubRecord record = new CommonObjectDataSubRecord(TestcaseRecordInputStream.Create((short)0x15, data),data.Length);


            Assert.AreEqual(CommonObjectType.ListBox, record.ObjectType);
            Assert.AreEqual((short)1, record.ObjectId);
            Assert.AreEqual((short)1, record.Option);
            Assert.AreEqual(true, record.IsLocked);
            Assert.AreEqual(false, record.IsPrintable);
            Assert.AreEqual(false, record.IsAutoFill);
            Assert.AreEqual(false, record.IsAutoline);
            Assert.AreEqual((int)24593, record.Reserved1);
            Assert.AreEqual((int)218103808, record.Reserved2);
            Assert.AreEqual((int)294, record.Reserved3);
            Assert.AreEqual(18, record.DataSize);
        }
        [Test]
        public void TestStore()
        {
            CommonObjectDataSubRecord record = new CommonObjectDataSubRecord();

            record.ObjectType = (CommonObjectType.ListBox);
            record.ObjectId = 1;
            record.Option = ((short)1);
            record.IsLocked = (true);
            record.IsPrintable = false;
            record.IsAutoFill = false;
            record.IsAutoline = false;
            record.Reserved1 = ((int)24593);
            record.Reserved2 = ((int)218103808);
            record.Reserved3 = ((int)294);

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}