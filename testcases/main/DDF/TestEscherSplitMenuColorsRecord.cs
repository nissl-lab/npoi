
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
    public class TestEscherSplitMenuColorsRecord
    {
        [Test]
        public void TestSerialize()
        {
            EscherSplitMenuColorsRecord r = CreateRecord();

            byte[] data = new byte[24];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(24, bytesWritten);
            Assert.AreEqual("[40, 00, " +
                    "1E, F1, " +
                    "10, 00, 00, 00, " +
                    "02, 04, 00, 00, " +
                    "02, 00, 00, 00, " +
                    "02, 00, 00, 00, " +
                    "01, 00, 00, 00]",
                    HexDump.ToHex(data));
        }
        [Test]
        public void TestFillFields()
        {
            String hexData = "40 00 " +
                    "1E F1 " +
                    "10 00 00 00 " +
                    "02 04 00 00 " +
                    "02 00 00 00 " +
                    "02 00 00 00 " +
                    "01 00 00 00 ";
            byte[] data = HexRead.ReadFromString(hexData);
            EscherSplitMenuColorsRecord r = new EscherSplitMenuColorsRecord();
            int bytesWritten = r.FillFields(data, new DefaultEscherRecordFactory());

            Assert.AreEqual(24, bytesWritten);
            Assert.AreEqual(0x0402, r.Color1);
            Assert.AreEqual(0x02, r.Color2);
            Assert.AreEqual(0x02, r.Color3);
            Assert.AreEqual(0x01, r.Color4);
        }
        [Test]
        public void TestToString()
        {
            String nl = Environment.NewLine;

            String expected = "EscherSplitMenuColorsRecord:" + nl +
                    "  RecordId: 0xF11E" + nl +
                    "  Version: 0x0000" + nl +
                    "  Instance: 0x0004" + nl +
                    "  Color1: 0x00000402" + nl +
                    "  Color2: 0x00000002" + nl +
                    "  Color3: 0x00000002" + nl +
                    "  Color4: 0x00000001" + nl +
                    "";
            Assert.AreEqual(expected, CreateRecord().ToString());
        }

        private EscherSplitMenuColorsRecord CreateRecord()
        {
            EscherSplitMenuColorsRecord r = new EscherSplitMenuColorsRecord();
            r.Options=(short)0x0040;
            r.RecordId=EscherSplitMenuColorsRecord.RECORD_ID;
            r.Color1=0x402;
            r.Color2=0x2;
            r.Color3=0x2;
            r.Color4=0x1;
            return r;
        }

    }
}