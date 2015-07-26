
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
    public class TestEscherDggRecord
    {
        [Test]
        public void TestSerialize()
        {
            EscherDggRecord r = CreateRecord();

            byte[] data = new byte[32];
            int bytesWritten = r.Serialize(0, data);
            Assert.AreEqual(32, bytesWritten);
            Assert.AreEqual("[00, 00, " +
                    "06, F0, " +
                    "18, 00, 00, 00, " +
                    "02, 04, 00, 00, " +
                    "02, 00, 00, 00, " +
                    "02, 00, 00, 00, " +
                    "01, 00, 00, 00, " +
                    "01, 00, 00, 00, 02, 00, 00, 00]",
                    HexDump.ToHex(data));
        }
        [Test]
        public void TestFillFields()
        {
            String hexData = "00 00 " +
                    "06 F0 " +
                    "18 00 00 00 " +
                    "02 04 00 00 " +
                    "02 00 00 00 " +
                    "02 00 00 00 " +
                    "01 00 00 00 " +
                    "01 00 00 00 02 00 00 00";
            byte[] data = HexRead.ReadFromString(hexData);
            EscherDggRecord r = new EscherDggRecord();
            int bytesWritten = r.FillFields(data, new DefaultEscherRecordFactory());

            Assert.AreEqual(32, bytesWritten);
            Assert.AreEqual(0x402, r.ShapeIdMax);
            Assert.AreEqual(0x02, r.NumIdClusters);
            Assert.AreEqual(0x02, r.NumShapesSaved);
            Assert.AreEqual(0x01, r.DrawingsSaved);
            Assert.AreEqual(1, r.FileIdClusters.Length);
            Assert.AreEqual(0x01, r.FileIdClusters[0].DrawingGroupId);
            Assert.AreEqual(0x02, r.FileIdClusters[0].NumShapeIdsUsed);
        }
        [Test]
        public void TestToString()
        {
            String nl = Environment.NewLine;

            String expected = "EscherDggRecord:" + nl +
                    "  RecordId: 0xF006" + nl +
                    "  Version: 0x0000" + nl +
                    "  Instance: 0x0000" + nl +
                    "  ShapeIdMax: 1026" + nl +
                    "  NumIdClusters: 2" + nl +
                    "  NumShapesSaved: 2" + nl +
                    "  DrawingsSaved: 1" + nl +
                    "  DrawingGroupId1: 1" + nl +
                    "  NumShapeIdsUsed1: 2" + nl;
            Assert.AreEqual(expected, CreateRecord().ToString());
        }

        private EscherDggRecord CreateRecord()
        {
            EscherDggRecord r = new EscherDggRecord();
            r.Options=(short)0x0000;
            r.RecordId=EscherDggRecord.RECORD_ID;
            r.ShapeIdMax=0x402;
            r.NumShapesSaved=0x02;
            r.DrawingsSaved=0x01;
            r.FileIdClusters=new EscherDggRecord.FileIdCluster[] {
            new EscherDggRecord.FileIdCluster( 1, 2 )
        };
            return r;
        }
        [Test]
        public void TestRecordSize()
        {
            EscherDggRecord r = new EscherDggRecord();
            r.FileIdClusters=new EscherDggRecord.FileIdCluster[] { new EscherDggRecord.FileIdCluster(0, 0) };
            Assert.AreEqual(32, r.RecordSize);

        }

    }
}