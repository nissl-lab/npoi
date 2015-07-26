
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
    public class TestEscherChildAnchorRecord
    {
        [Test]
        public void TestSerialize()
        {
            EscherChildAnchorRecord r = CreateRecord();

            byte[] data = new byte[8 + 16];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(24, bytesWritten);
            Assert.AreEqual("[01, 00, " +
                    "0F, F0, " +
                    "10, 00, 00, 00, " +
                    "01, 00, 00, 00, " +
                    "02, 00, 00, 00, " +
                    "03, 00, 00, 00, " +
                    "04, 00, 00, 00]", HexDump.ToHex(data));
        }
        [Test]
        public void TestFillFields()
        {
            String hexData = "01 00 " +
                    "0F F0 " +
                    "10 00 00 00 " +
                    "01 00 00 00 " +
                    "02 00 00 00 " +
                    "03 00 00 00 " +
                    "04 00 00 00 ";

            byte[] data = HexRead.ReadFromString(hexData);
            EscherChildAnchorRecord r = new EscherChildAnchorRecord();
            int bytesWritten = r.FillFields(data, new DefaultEscherRecordFactory());

            Assert.AreEqual(24, bytesWritten);
            Assert.AreEqual(1, r.Dx1);
            Assert.AreEqual(2, r.Dy1);
            Assert.AreEqual(3, r.Dx2);
            Assert.AreEqual(4, r.Dy2);
            Assert.AreEqual((short)0x0001, r.Options);
        }
        [Test]
        public void TestToString()
        {
            String nl = Environment.NewLine;

            String expected = "EscherChildAnchorRecord:" + nl +
                    "  RecordId: 0xF00F" + nl +
                    "  Version: 0x0001" + nl +
                    "  Instance: 0x0000" + nl +
                    "  X1: 1" + nl +
                    "  Y1: 2" + nl +
                    "  X2: 3" + nl +
                    "  Y2: 4" + nl;
            Assert.AreEqual(expected, CreateRecord().ToString());
        }

        private EscherChildAnchorRecord CreateRecord()
        {
            EscherChildAnchorRecord r = new EscherChildAnchorRecord();
            r.RecordId=EscherChildAnchorRecord.RECORD_ID;
            r.Options=(short)0x0001;
            r.Dx1=1;
            r.Dy1=2;
            r.Dx2=3;
            r.Dy2=4;
            return r;
        }

    }
}