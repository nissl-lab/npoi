
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
    public class TestEscherBlipWMFRecord
    {
        private String dataStr;
        private byte[] data;
        
        public TestEscherBlipWMFRecord()
        {
            dataStr = "2C 15 18 F0 34 00 00 00 01 01 01 01 01 01 01 01 " +
                            "01 01 01 01 01 01 01 01 06 00 00 00 03 00 00 00 " +
                            "01 00 00 00 04 00 00 00 02 00 00 00 0A 00 00 00 " +
                            "0B 00 00 00 05 00 00 00 08 07 01 02";
            data = HexRead.ReadFromString(dataStr);
        }
        [Test]
        public void TestSerialize()
        {
            EscherBlipWMFRecord r = new EscherBlipWMFRecord();
            r.BoundaryLeft=1;
            r.BoundaryHeight=2;
            r.BoundaryTop=3;
            r.BoundaryWidth=4;
            r.CacheOfSavedSize=5;
            r.CacheOfSize=6;
            r.Filter=(byte)7;
            r.CompressionFlag=(byte)8;
            r.SecondaryUID=new byte[] { (byte)0x01, (byte)0x01, (byte)0x01, (byte)0x01,
                                       (byte)0x01, (byte)0x01, (byte)0x01, (byte)0x01,
                                       (byte)0x01, (byte)0x01, (byte)0x01, (byte)0x01,
                                       (byte)0x01, (byte)0x01, (byte)0x01, (byte)0x01,  };
            r.Width=10;
            r.Height=11;
            r.RecordId=EscherBlipWMFRecord.RECORD_ID_START;
            r.Options=(short)5420;
            r.Data=new byte[] { (byte)0x01, (byte)0x02 };

            byte[] buf = new byte[r.RecordSize];
            r.Serialize(0, buf);

            Assert.AreEqual("[2C, 15, 18, F0, 26, 00, 00, 00, " +
                    "01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, " +
                    "06, 00, 00, 00, " +    // field_2_cacheOfSize
                    "03, 00, 00, 00, " +    // field_3_boundaryTop
                    "01, 00, 00, 00, " +    // field_4_boundaryLeft
                    "04, 00, 00, 00, " +    // field_5_boundaryWidth
                    "02, 00, 00, 00, " +    // field_6_boundaryHeight
                    "0A, 00, 00, 00, " +    // field_7_x
                    "0B, 00, 00, 00, " +    // field_8_y
                    "05, 00, 00, 00, " +    // field_9_cacheOfSavedSize
                    "08, " +                // field_10_compressionFlag
                    "07, " +                // field_11_filter
                    "01, 02]",            // field_12_data
                    HexDump.ToHex(buf));
            Assert.AreEqual(60, r.RecordSize);

        }
        [Test]
        public void TestFillFields()
        {
            EscherBlipWMFRecord r = new EscherBlipWMFRecord();
            r.FillFields(data, 0, new DefaultEscherRecordFactory());

            Assert.AreEqual(EscherBlipWMFRecord.RECORD_ID_START, r.RecordId);
            Assert.AreEqual(1, r.BoundaryLeft);
            Assert.AreEqual(2, r.BoundaryHeight);
            Assert.AreEqual(3, r.BoundaryTop);
            Assert.AreEqual(4, r.BoundaryWidth);
            Assert.AreEqual(5, r.CacheOfSavedSize);
            Assert.AreEqual(6, r.CacheOfSize);
            Assert.AreEqual(7, r.Filter);
            Assert.AreEqual(8, r.CompressionFlag);
            Assert.AreEqual("[01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01]", HexDump.ToHex(r.SecondaryUID));
            Assert.AreEqual(10, r.Width);
            Assert.AreEqual(11, r.Height);
            Assert.AreEqual((short)5420, r.Options);
            Assert.AreEqual("[01, 02]", HexDump.ToHex(r.Data));
        }
        [Test]
        public void TestToString()
        {
            EscherBlipWMFRecord r = new EscherBlipWMFRecord();
            r.FillFields(data, 0, new DefaultEscherRecordFactory());

            String nl = Environment.NewLine;

            Assert.AreEqual("EscherBlipWMFRecord:" + nl +
                    "  RecordId: 0xF018" + nl +
                    "  Version: 0x000C" + nl +
                    "  Instance: 0x0152" + nl +
                    "  Secondary UID: [01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01, 01]" + nl +
                    "  CacheOfSize: 6" + nl +
                    "  BoundaryTop: 3" + nl +
                    "  BoundaryLeft: 1" + nl +
                    "  BoundaryWidth: 4" + nl +
                    "  BoundaryHeight: 2" + nl +
                    "  X: 10" + nl +
                    "  Y: 11" + nl +
                    "  CacheOfSavedSize: 5" + nl +
                    "  CompressionFlag: 8" + nl +
                    "  Filter: 7" + nl +
                    "  Data:" + nl +
                    "00000000 01 02                                           .." + nl
                    , r.ToString());
        }

    }
}
