
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

namespace TestCases.DDF
{

    using System;
    using System.Text;
    using System.Collections.Generic;
    using System.IO;

    using NUnit.Framework;
    using NPOI.DDF;
    using NPOI.Util;
    [TestFixture]
    public class TestUnknownEscherRecord
    {
        [Test]
        public void TestFillFields()
        {
            String TestData =
                    "0F 02 " + // options
                    "11 F1 " + // record id
                    "00 00 00 00";      // remaining bytes

            UnknownEscherRecord r = new UnknownEscherRecord();
            IEscherRecordFactory factory = new DefaultEscherRecordFactory();
            r.FillFields(HexRead.ReadFromString(TestData), factory);

            Assert.AreEqual(0x020F, r.Options);
            Assert.AreEqual(unchecked((short)0xF111), r.RecordId);
            Assert.IsTrue(r.IsContainerRecord);
            Assert.AreEqual(8, r.RecordSize);
            Assert.AreEqual(0, r.ChildRecords.Count);
            Assert.AreEqual(0, r.Data.Length);

            TestData =
                    "00 02 " + // options
                    "11 F1 " + // record id
                    "04 00 00 00 " + // remaining bytes
                    "01 02 03 04";

            r = new UnknownEscherRecord();
            r.FillFields(HexRead.ReadFromString(TestData), factory);

            Assert.AreEqual(0x0200, r.Options);
            Assert.AreEqual(unchecked((short)0xF111), r.RecordId);
            Assert.AreEqual(12, r.RecordSize);
            Assert.IsFalse(r.IsContainerRecord);
            Assert.AreEqual(0, r.ChildRecords.Count);
            Assert.AreEqual(4, r.Data.Length);
            Assert.AreEqual(1, r.Data[0]);
            Assert.AreEqual(2, r.Data[1]);
            Assert.AreEqual(3, r.Data[2]);
            Assert.AreEqual(4, r.Data[3]);

            TestData =
                    "0F 02 " + // options
                    "11 F1 " + // record id
                    "08 00 00 00 " + // remaining bytes
                    "00 02 " + // options
                    "FF FF " + // record id
                    "00 00 00 00";      // remaining bytes

            r = new UnknownEscherRecord();
            r.FillFields(HexRead.ReadFromString(TestData), factory);

            Assert.AreEqual(0x020F, r.Options);
            Assert.AreEqual(unchecked((short)0xF111), r.RecordId);
            Assert.AreEqual(8, r.RecordSize);
            Assert.IsTrue(r.IsContainerRecord);
            Assert.AreEqual(1, r.ChildRecords.Count);
            Assert.AreEqual(unchecked((short)0xFFFF), r.GetChild(0).RecordId);

        }
        [Test]
        public void TestSerialize()
        {
            UnknownEscherRecord r = new UnknownEscherRecord();
            r.Options=(short)0x1234;
            r.RecordId=unchecked((short)0xF112);
            byte[] data = new byte[8];
            r.Serialize(0, data);

            Assert.AreEqual("[34, 12, 12, F1, 00, 00, 00, 00]", HexDump.ToHex(data));

            EscherRecord childRecord = new UnknownEscherRecord();
            childRecord.Options=unchecked((short)0x9999);
            childRecord.RecordId=unchecked((short)0xFF01);
            r.AddChildRecord(childRecord);
            r.Options=(short)0x123F;
            data = new byte[16];
            r.Serialize(0, data);

            Assert.AreEqual("[3F, 12, 12, F1, 08, 00, 00, 00, 99, 99, 01, FF, 00, 00, 00, 00]", HexDump.ToHex(data));
        }
        [Test]
        public void TestToString()
        {
            UnknownEscherRecord r = new UnknownEscherRecord();
            r.Options=(short)0x1234;
            r.RecordId=unchecked((short)0xF112);
            byte[] data = new byte[8];
            r.Serialize(0, data);

            String nl = Environment.NewLine;
            Assert.AreEqual("UnknownEscherRecord:" + nl +
                    "  isContainer: False" + nl +
                    "  version: 0x0004" + nl +
                    "  instance: 0x0123" + nl +
                    "  recordId: 0xF112" + nl +
                    "  numchildren: 0" + nl
                    , r.ToString());
        }


    }
}
