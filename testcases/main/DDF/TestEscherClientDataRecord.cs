
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
    public class TestEscherClientDataRecord
    {
        [Test]
        public void TestSerialize()
        {
            EscherClientDataRecord r = CreateRecord();

            byte[] data = new byte[8];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(8, bytesWritten);
            Assert.AreEqual("[02, 00, " +
                    "11, F0, " +
                    "00, 00, 00, 00]",
                    HexDump.ToHex(data));
        }
        [Test]
        public void TestFillFields()
        {
            String hexData = "02 00 " +
                    "11 F0 " +
                    "00 00 00 00 ";
            byte[] data = HexRead.ReadFromString(hexData);
            EscherClientDataRecord r = new EscherClientDataRecord();
            int bytesWritten = r.FillFields(data, new DefaultEscherRecordFactory());

            Assert.AreEqual(8, bytesWritten);
            Assert.AreEqual(unchecked((short)0xF011), r.RecordId);
            Assert.AreEqual("[]", HexDump.ToHex(r.RemainingData));
        }
        [Test]
        public void TestToString()
        {
            String nl = Environment.NewLine;

            String expected = "EscherClientDataRecord:" + nl +
                    "  RecordId: 0xF011" + nl +
                    "  Version: 0x0002" + nl +
                    "  Instance: 0x0000" + nl +
                    "  Extra Data:" + nl +
                    "No Data" + nl;
            Assert.AreEqual(expected, CreateRecord().ToString());
        }

        private EscherClientDataRecord CreateRecord()
        {
            EscherClientDataRecord r = new EscherClientDataRecord();
            r.Options=(short)0x0002;
            r.RecordId=EscherClientDataRecord.RECORD_ID;
            r.RemainingData=new byte[] { };
            return r;
        }

    }
}