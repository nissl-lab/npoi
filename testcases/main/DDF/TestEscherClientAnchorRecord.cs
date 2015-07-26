
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
    public class TestEscherClientAnchorRecord
    {
        [Test]
        public void TestSerialize()
        {
            EscherClientAnchorRecord r = CreateRecord();

            byte[] data = new byte[8 + 18 + 2];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(28, bytesWritten);
            Assert.AreEqual("[01, 00, " +
                    "10, F0, " +
                    "14, 00, 00, 00, " +
                    "4D, 00, 37, 00, 21, 00, 58, 00, " +
                    "0B, 00, 2C, 00, 16, 00, 63, 00, " +
                    "42, 00, " +
                    "FF, DD]", HexDump.ToHex(data));
        }
        [Test]
        public void TestFillFields()
        {
            String hexData = "01 00 " +
                    "10 F0 " +
                    "14 00 00 00 " +
                    "4D 00 37 00 21 00 58 00 " +
                    "0B 00 2C 00 16 00 63 00 " +
                    "42 00 " +
                    "FF DD";
            byte[] data = HexRead.ReadFromString(hexData);
            EscherClientAnchorRecord r = new EscherClientAnchorRecord();
            int bytesWritten = r.FillFields(data, new DefaultEscherRecordFactory());

            Assert.AreEqual(28, bytesWritten);
            Assert.AreEqual((short)55, r.Col1);
            Assert.AreEqual((short)44, r.Col2);
            Assert.AreEqual((short)33, r.Dx1);
            Assert.AreEqual((short)22, r.Dx2);
            Assert.AreEqual((short)11, r.Dy1);
            Assert.AreEqual((short)66, r.Dy2);
            Assert.AreEqual((short)77, r.Flag);
            Assert.AreEqual((short)88, r.Row1);
            Assert.AreEqual((short)99, r.Row2);
            Assert.AreEqual((short)0x0001, r.Options);
            Assert.AreEqual((byte)0xFF, r.RemainingData[0]);
            Assert.AreEqual((byte)0xDD, r.RemainingData[1]);
        }
        [Test]
        public void TestToString()
        {
            String nl = Environment.NewLine;

            String expected = "EscherClientAnchorRecord:" + nl +
                    "  RecordId: 0xF010" + nl +
                    "  Version: 0x0001" + nl +
                    "  Instance: 0x0000" + nl +
                    "  Flag: 77" + nl +
                    "  Col1: 55" + nl +
                    "  DX1: 33" + nl +
                    "  Row1: 88" + nl +
                    "  DY1: 11" + nl +
                    "  Col2: 44" + nl +
                    "  DX2: 22" + nl +
                    "  Row2: 99" + nl +
                    "  DY2: 66" + nl +
                    "  Extra Data:" + nl +
                    "00000000 FF DD                                           .." + nl;
            Assert.AreEqual(expected, CreateRecord().ToString());
        }

        private EscherClientAnchorRecord CreateRecord()
        {
            EscherClientAnchorRecord r = new EscherClientAnchorRecord();
            r.Col1=(short)55;
            r.Col2=(short)44;
            r.Dx1=(short)33;
            r.Dx2=(short)22;
            r.Dy1=(short)11;
            r.Dy2=(short)66;
            r.Flag=(short)77;
            r.Row1=(short)88;
            r.Row2=(short)99;
            r.Options=(short)0x0001;
            r.RemainingData=new byte[] { (byte)0xFF, (byte)0xDD };
            return r;
        }

    }
}