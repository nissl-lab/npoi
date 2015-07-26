
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
    public class TestEscherDgRecord
    {
        [Test]
        public void TestSerialize()
        {
            EscherDgRecord r = CreateRecord();

            byte[] data = new byte[16];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(16, bytesWritten);
            Assert.AreEqual("[10, 00, " +
                    "08, F0, " +
                    "08, 00, 00, 00, " +
                    "02, 00, 00, 00, " +     // num shapes in drawing
                    "01, 04, 00, 00]",     // The last MSOSPID given to an SP in this DG
                    HexDump.ToHex(data));
        }
        [Test]
        public void TestFillFields()
        {
            String hexData = "10 00 " +
                    "08 F0 " +
                    "08 00 00 00 " +
                    "02 00 00 00 " +
                    "01 04 00 00 ";
            byte[] data = HexRead.ReadFromString(hexData);
            EscherDgRecord r = new EscherDgRecord();
            int bytesWritten = r.FillFields(data, new DefaultEscherRecordFactory());

            Assert.AreEqual(16, bytesWritten);
            Assert.AreEqual(2, r.NumShapes);
            Assert.AreEqual(1025, r.LastMSOSPID);
        }
        [Test]
        public void TestToString()
        {
            String nl = Environment.NewLine;

            String expected = "EscherDgRecord:" + nl +
                    "  RecordId: 0xF008" +nl +    
                    "  Version: 0x0000" + nl +
                    "  Instance: 0x0001" + nl +
                    "  NumShapes: 2" + nl +
                    "  LastMSOSPID: 1025" + nl;
            Assert.AreEqual(expected, CreateRecord().ToString());
        }

        private EscherDgRecord CreateRecord()
        {
            EscherDgRecord r = new EscherDgRecord();
            r.Options=(short)0x0010;
            r.RecordId=EscherDgRecord.RECORD_ID;
            r.NumShapes=2;
            r.LastMSOSPID=1025;
            return r;
        }

    }
}