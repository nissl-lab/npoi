/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is1 distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace TestCases.HSSF.Record
{
    using System;
    using NPOI.HSSF.Record;
    using NUnit.Framework;
    using NPOI.DDF;
    using NPOI.Util;

    [TestFixture]
    public class TestDrawingGroupRecord
    {
        static int MAX_RECORD_SIZE = 8228;
        private static int MAX_DATA_SIZE = MAX_RECORD_SIZE - 4;

        [Test]
        public void TestRecordSize()
        {
            DrawingGroupRecord r = new DrawingGroupRecord();
            Assert.AreEqual(4, r.RecordSize);

            EscherSpRecord sp = new EscherSpRecord();
            sp.RecordId = (EscherSpRecord.RECORD_ID);
            sp.Options = ((short)0x1111);
            sp.Flags = (-1);
            sp.ShapeId = (-1);
            EscherContainerRecord dggContainer = new EscherContainerRecord();
            dggContainer.Options = ((short)0x000F);
            dggContainer.RecordId = unchecked((short)0xF000);
            dggContainer.AddChildRecord(sp);

            r.AddEscherRecord(dggContainer);
            Assert.AreEqual(28, r.RecordSize);

            byte[] data = new byte[28];
            int size = r.Serialize(0, data);
            Assert.AreEqual("[EB, 00, 18, 00, 0F, 00, 00, F0, 10, 00, 00, 00, 11, 11, 0A, F0, 08, 00, 00, 00, FF, FF, FF, FF, FF, FF, FF, FF]", HexDump.ToHex(data));
            Assert.AreEqual(28, size);

            Assert.AreEqual(24, dggContainer.RecordSize);


            r = new DrawingGroupRecord();
            r.RawData = (new byte[MAX_DATA_SIZE]);
            Assert.AreEqual(MAX_RECORD_SIZE, r.RecordSize);
            r.RawData = (new byte[MAX_DATA_SIZE + 1]);
            Assert.AreEqual(MAX_RECORD_SIZE + 5, r.RecordSize);
            r.RawData = (new byte[MAX_DATA_SIZE * 2]);
            Assert.AreEqual(MAX_RECORD_SIZE * 2, r.RecordSize);
            r.RawData = (new byte[MAX_DATA_SIZE * 2 + 1]);
            Assert.AreEqual(MAX_RECORD_SIZE * 2 + 5, r.RecordSize);
        }
        [Test]
        public void TestSerialize()
        {
            // Check under max record size
            DrawingGroupRecord r = new DrawingGroupRecord();
            byte[] rawData = new byte[100];
            rawData[0] = 100;
            rawData[99] = (byte)200;
            r.RawData = (rawData);
            byte[] buffer = new byte[r.RecordSize];
            int size = r.Serialize(0, buffer);
            Assert.AreEqual(104, size);
            Assert.AreEqual("[EB, 00, 64, 00, 64, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, C8]", HexDump.ToHex(buffer));

            // check at max record size
            rawData = new byte[MAX_DATA_SIZE];
            r.RawData = (rawData);
            buffer = new byte[r.RecordSize];
            size = r.Serialize(0, buffer);
            Assert.AreEqual(MAX_RECORD_SIZE, size);

            // check over max record size
            rawData = new byte[MAX_DATA_SIZE + 1];
            rawData[rawData.Length - 1] = (byte)255;
            r.RawData = (rawData);
            buffer = new byte[r.RecordSize];
            size = r.Serialize(0, buffer);
            Assert.AreEqual(MAX_RECORD_SIZE + 5, size);
            Assert.AreEqual("[EB, 00, 20, 20]", HexDump.ToHex(cut(buffer, 0, 4)));
            Assert.AreEqual("[00, EB, 00, 01, 00, FF]", HexDump.ToHex(cut(buffer, MAX_RECORD_SIZE - 1, MAX_RECORD_SIZE + 5)));

            // check continue record
            rawData = new byte[MAX_DATA_SIZE * 2 + 1];
            rawData[rawData.Length - 1] = (byte)255;
            r.RawData = (rawData);
            buffer = new byte[r.RecordSize];
            size = r.Serialize(0, buffer);
            Assert.AreEqual(MAX_RECORD_SIZE * 2 + 5, size);
            Assert.AreEqual(MAX_RECORD_SIZE * 2 + 5, r.RecordSize);
            Assert.AreEqual("[EB, 00, 20, 20]", HexDump.ToHex(cut(buffer, 0, 4)));
            Assert.AreEqual("[EB, 00, 20, 20]", HexDump.ToHex(cut(buffer, MAX_RECORD_SIZE, MAX_RECORD_SIZE + 4)));
            Assert.AreEqual("[3C, 00, 01, 00, FF]", HexDump.ToHex(cut(buffer, MAX_RECORD_SIZE * 2, MAX_RECORD_SIZE * 2 + 5)));

            // check continue record
            rawData = new byte[664532];
            r.RawData = (rawData);
            buffer = new byte[r.RecordSize];
            size = r.Serialize(0, buffer);
            Assert.AreEqual(664856, size);
            Assert.AreEqual(664856, r.RecordSize);
        }

        private byte[] cut(byte[] data, int fromInclusive, int toExclusive)
        {
            int length = toExclusive - fromInclusive;
            byte[] result = new byte[length];
            System.Array.Copy(data, fromInclusive, result, 0, length);
            return result;
        }
        [Test]
        public void TestGrossSizeFromDataSize()
        {
            for (int i = 0; i < MAX_RECORD_SIZE * 4; i += 11)
            {
                //System.out.print( "data size = " + i + ", gross size = " + DrawingGroupRecord.GrossSizeFromDataSize( i ) );
                //System.out.println( "  Diff: " + (DrawingGroupRecord.GrossSizeFromDataSize( i ) - i) );
            }

            Assert.AreEqual(4, DrawingGroupRecord.GrossSizeFromDataSize(0));
            Assert.AreEqual(5, DrawingGroupRecord.GrossSizeFromDataSize(1));
            Assert.AreEqual(MAX_RECORD_SIZE, DrawingGroupRecord.GrossSizeFromDataSize(MAX_DATA_SIZE));
            Assert.AreEqual(MAX_RECORD_SIZE + 5, DrawingGroupRecord.GrossSizeFromDataSize(MAX_DATA_SIZE + 1));
            Assert.AreEqual(MAX_RECORD_SIZE * 2, DrawingGroupRecord.GrossSizeFromDataSize(MAX_DATA_SIZE * 2));
            Assert.AreEqual(MAX_RECORD_SIZE * 2 + 5, DrawingGroupRecord.GrossSizeFromDataSize(MAX_DATA_SIZE * 2 + 1));
        }

    }
}
