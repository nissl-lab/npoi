
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
    using System.IO;

    using NUnit.Framework;
    using NPOI.DDF;
    using NPOI.Util;
    [TestFixture]
    public class TestEscherSpgrRecord
    {
        [Test]
        public void TestSerialize()
        {
            EscherSpgrRecord r = CreateRecord();

            byte[] data = new byte[24];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(24, bytesWritten);
            Assert.AreEqual("[10, 00, " +
                    "09, F0, " +
                    "10, 00, 00, 00, " +
                    "01, 00, 00, 00, " +     // x
                    "02, 00, 00, 00, " +     // y
                    "03, 00, 00, 00, " +     // width
                    "04, 00, 00, 00]",     // height
                    HexDump.ToHex(data));
        }
        [Test]
        public void TestFillFields()
        {
            String hexData = "10 00 " +
                    "09 F0 " +
                    "10 00 00 00 " +
                    "01 00 00 00 " +
                    "02 00 00 00 " +
                    "03 00 00 00 " +
                    "04 00 00 00 ";
            byte[] data = HexRead.ReadFromString(hexData);
            EscherSpgrRecord r = new EscherSpgrRecord();
            int bytesWritten = r.FillFields(data, new DefaultEscherRecordFactory());

            Assert.AreEqual(24, bytesWritten);
            Assert.AreEqual(1, r.RectX1);
            Assert.AreEqual(2, r.RectY1);
            Assert.AreEqual(3, r.RectX2);
            Assert.AreEqual(4, r.RectY2);
        }
        [Test]
        public void TestToString()
        {
            String nl = Environment.NewLine;

            String expected = "EscherSpgrRecord:" + nl +
                    "  RecordId: 0xF009" + nl +
                    "  Version: 0x0000" + nl +
                    "  Instance: 0x0001" + nl +
                    "  RectX: 1" + nl +
                    "  RectY: 2" + nl +
                    "  RectWidth: 3" + nl +
                    "  RectHeight: 4" + nl;
            ;
            Assert.AreEqual(expected, CreateRecord().ToString());
        }

        private EscherSpgrRecord CreateRecord()
        {
            EscherSpgrRecord r = new EscherSpgrRecord();
            r.Options=(short)0x0010;
            r.RecordId=EscherSpgrRecord.RECORD_ID;
            r.RectX1=1;
            r.RectY1=2;
            r.RectX2=3;
            r.RectY2=4;
            return r;
        }

    }
}