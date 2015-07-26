
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
    public class TestEscherBSERecord
    {
        [Test]
        public void TestFillFields()
        {
            String data = "01 00 00 00 24 00 00 00 05 05 01 02 03 04 " +
                    " 05 06 07 08 09 0A 0B 0C 0D 0E 0F 00 01 00 00 00 " +
                    " 00 00 02 00 00 00 03 00 00 00 04 05 06 07";
            EscherBSERecord r = new EscherBSERecord();
            int bytesWritten = r.FillFields(HexRead.ReadFromString(data), 0, new DefaultEscherRecordFactory());
            Assert.AreEqual(44, bytesWritten);
            Assert.AreEqual((short)0x0001, r.Options);
            Assert.AreEqual(EscherBSERecord.BT_JPEG, r.BlipTypeWin32);
            Assert.AreEqual(EscherBSERecord.BT_JPEG, r.BlipTypeMacOS);
            Assert.AreEqual("[01, 02, 03, 04, 05, 06, 07, 08, 09, 0A, 0B, 0C, 0D, 0E, 0F, 00]", HexDump.ToHex(r.UID));
            Assert.AreEqual((short)1, r.Tag);
            Assert.AreEqual(2, r.Ref);
            Assert.AreEqual(3, r.Offset);
            Assert.AreEqual((byte)4, r.Usage);
            Assert.AreEqual((byte)5, r.Name);
            Assert.AreEqual((byte)6, r.Unused2);
            Assert.AreEqual((byte)7, r.Unused3);
            Assert.AreEqual(0, r.RemainingData.Length);
        }
        [Test]
        public void TestSerialize()
        {
            EscherBSERecord r = CreateRecord();

            byte[] data = new byte[8 + 36];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(44, bytesWritten);
            Assert.AreEqual("[01, 00, 00, 00, 24, 00, 00, 00, 05, 05, 01, 02, 03, 04, " +
                    "05, 06, 07, 08, 09, 0A, 0B, 0C, 0D, 0E, 0F, 00, 01, 00, 00, 00, " +
                    "00, 00, 02, 00, 00, 00, 03, 00, 00, 00, 04, 05, 06, 07]",
                    HexDump.ToHex(data));

        }

        private EscherBSERecord CreateRecord()
        {
            EscherBSERecord r = new EscherBSERecord();
            r.Options=(short)0x0001;
            r.BlipTypeWin32=EscherBSERecord.BT_JPEG;
            r.BlipTypeMacOS=EscherBSERecord.BT_JPEG;
            r.UID=HexRead.ReadFromString("01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F 00");
            r.Tag=(short)1;
            r.Ref=2;
            r.Offset=3;
            r.Usage=(byte)4;
            r.Name=(byte)5;
            r.Unused2=(byte)6;
            r.Unused3=(byte)7;
            r.RemainingData=new byte[0];
            return r;

        }
        [Test]
        public void TestToString()
        {
            EscherBSERecord record = CreateRecord();
            String nl = Environment.NewLine;
            Assert.AreEqual("EscherBSERecord:" + nl +
                    "  RecordId: 0xF007" + nl +
                    "  Version: 0x0001" + '\n' +
                    "  Instance: 0x0000" + '\n' +
                    "  BlipTypeWin32: 5" + nl +
                    "  BlipTypeMacOS: 5" + nl +
                    "  SUID: [01, 02, 03, 04, 05, 06, 07, 08, 09, 0A, 0B, 0C, 0D, 0E, 0F, 00]" + nl +
                    "  Tag: 1" + nl +
                    "  Size: 0" + nl +
                    "  Ref: 2" + nl +
                    "  Offset: 3" + nl +
                    "  Usage: 4" + nl +
                    "  Name: 5" + nl +
                    "  Unused2: 6" + nl +
                    "  Unused3: 7" + nl +
                    "  blipRecord: null" + nl +
                    "  Extra Data:" + nl +
                    "No Data" + nl
                    , record.ToString());
        }

    }
}