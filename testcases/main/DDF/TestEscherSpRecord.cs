
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
    public class TestEscherSpRecord
    {
        [Test]
        public void TestSerialize()
        {
            EscherSpRecord r = CreateRecord();

            byte[] data = new byte[16];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(16, bytesWritten);
            Assert.AreEqual("[02, 00, " +
                    "0A, F0, " +
                    "08, 00, 00, 00, " +
                    "00, 04, 00, 00, " +
                    "05, 00, 00, 00]",
                    HexDump.ToHex(data));
        }
        [Test]
        public void TestFillFields()
        {
            String hexData = "02 00 " +
                    "0A F0 " +
                    "08 00 00 00 " +
                    "00 04 00 00 " +
                    "05 00 00 00 ";
            byte[] data = HexRead.ReadFromString(hexData);
            EscherSpRecord r = new EscherSpRecord();
            int bytesWritten = r.FillFields(data, new DefaultEscherRecordFactory());

            Assert.AreEqual(16, bytesWritten);
            Assert.AreEqual(0x0400, r.ShapeId);
            Assert.AreEqual(0x05, r.Flags);
        }
        [Test]
        public void TestToString()
        {
            String nl = Environment.NewLine;

            String expected = "EscherSpRecord:" + nl +
                    "  RecordId: 0xF00A" + nl +
                    "  Version: 0x0002" + nl +
                    "  ShapeType: 0x0000" + nl +
                    "  ShapeId: 1024" + nl +
                    "  Flags: GROUP|PATRIARCH (0x00000005)" + nl;
            Assert.AreEqual(expected, CreateRecord().ToString());
        }

        private EscherSpRecord CreateRecord()
        {
            EscherSpRecord r = new EscherSpRecord();
            r.Options=(short)0x0002;
            r.RecordId=EscherSpRecord.RECORD_ID;
            r.ShapeId=0x0400;
            r.Flags=0x05;
            return r;
        }

    }
}