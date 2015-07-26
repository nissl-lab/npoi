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
    using System.Collections;

    using NPOI.DDF;
    using NPOI.Util;
    using NPOI.HSSF.Model;
    using NPOI.HSSF.Record;
    using NUnit.Framework;
    using System.Collections.Generic;

    /**
     * Tests the EscherAggregate class.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestEscherAggregate
    {
        /**
         * Tests that the create aggregate method correctly rejoins escher records toGether.
         *
         * @
         */
        [Test]
        public void TestCreateAggregate()
        {
            String msoDrawingRecord1 =
                    "0F 00 02 F0 20 01 00 00 10 00 08 F0 08 00 00 00 " +
                    "03 00 00 00 02 04 00 00 0F 00 03 F0 08 01 00 00 " +
                    "0F 00 04 F0 28 00 00 00 01 00 09 F0 10 00 00 00 " +
                    "00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 " +
                    "02 00 0A F0 08 00 00 00 00 04 00 00 05 00 00 00 " +
                    "0F 00 04 F0 64 00 00 00 42 01 0A F0 08 00 00 00 " +
                    "01 04 00 00 00 0A 00 00 73 00 0B F0 2A 00 00 00 " +
                    "BF 00 08 00 08 00 44 01 04 00 00 00 7F 01 00 00 " +
                    "01 00 BF 01 00 00 11 00 C0 01 40 00 00 08 FF 01 " +
                    "10 00 10 00 BF 03 00 00 08 00 00 00 10 F0 12 00 " +
                    "00 00 00 00 01 00 54 00 05 00 45 00 01 00 88 03 " +
                    "05 00 94 00 00 00 11 F0 00 00 00 00";

            String msoDrawingRecord2 =
                    "0F 00 04 F0 64 00 00 00 42 01 0A F0 08 00 00 00 " +
                    "02 04 00 00 80 0A 00 00 73 00 0B F0 2A 00 00 00 " +
                    "BF 00 08 00 08 00 44 01 04 00 00 00 7F 01 00 00 " +
                    "01 00 BF 01 00 00 11 00 C0 01 40 00 00 08 FF 01 " +
                    "10 00 10 00 BF 03 00 00 08 00 00 00 10 F0 12 00 " +
                    "00 00 00 00 01 00 8D 03 05 00 E4 00 03 00 4D 03 " +
                    "0B 00 0C 00 00 00 11 F0 00 00 00 00";

            DrawingRecord d1 = new DrawingRecord();
            d1.Data = (HexRead.ReadFromString(msoDrawingRecord1));

            ObjRecord r1 = new ObjRecord();

            DrawingRecord d2 = new DrawingRecord();
            d2.Data = (HexRead.ReadFromString(msoDrawingRecord2));

            ObjRecord r2 = new ObjRecord();

            List<RecordBase> records = new List<RecordBase>();
            records.Add(d1);
            records.Add(r1);
            records.Add(d2);
            records.Add(r2);

            EscherAggregate aggregate = EscherAggregate.CreateAggregate(records, 0);

            Assert.AreEqual(1, aggregate.EscherRecords.Count);
            Assert.AreEqual(unchecked((short)0xF002), aggregate.GetEscherRecord(0).RecordId);
            Assert.AreEqual(2, aggregate.GetEscherRecord(0).ChildRecords.Count);

            //        System.out.println( "aggregate = " + aggregate );
        }
        [Test]
        public void TestSerialize()
        {

            EscherContainerRecord container1 = new EscherContainerRecord();
            EscherContainerRecord spContainer1 = new EscherContainerRecord();
            EscherContainerRecord spContainer2 = new EscherContainerRecord();
            EscherContainerRecord spContainer3 = new EscherContainerRecord();
            EscherSpRecord sp1 = new EscherSpRecord();
            EscherSpRecord sp2 = new EscherSpRecord();
            EscherSpRecord sp3 = new EscherSpRecord();
            EscherClientDataRecord d2 = new EscherClientDataRecord();
            EscherClientDataRecord d3 = new EscherClientDataRecord();

            container1.Options = ((short)0x000F);
            spContainer1.Options = ((short)0x000F);
            spContainer1.RecordId = (EscherContainerRecord.SP_CONTAINER);
            spContainer2.Options = ((short)0x000F);
            spContainer2.RecordId = (EscherContainerRecord.SP_CONTAINER);
            spContainer3.Options = ((short)0x000F);
            spContainer3.RecordId = (EscherContainerRecord.SP_CONTAINER);
            d2.RecordId = (EscherClientDataRecord.RECORD_ID);
            d2.RemainingData = (new byte[0]);
            d3.RecordId = (EscherClientDataRecord.RECORD_ID);
            d3.RemainingData = (new byte[0]);
            container1.AddChildRecord(spContainer1);
            container1.AddChildRecord(spContainer2);
            container1.AddChildRecord(spContainer3);
            spContainer1.AddChildRecord(sp1);
            spContainer2.AddChildRecord(sp2);
            spContainer3.AddChildRecord(sp3);
            spContainer2.AddChildRecord(d2);
            spContainer3.AddChildRecord(d3);

            EscherAggregate aggregate = new EscherAggregate(false);
            aggregate.AddEscherRecord(container1);
            aggregate.AssociateShapeToObjRecord(d2, new ObjRecord());
            aggregate.AssociateShapeToObjRecord(d3, new ObjRecord());

            byte[] data = new byte[112];
            int bytesWritten = aggregate.Serialize(0, data);
            Assert.AreEqual(112, bytesWritten);
            Assert.AreEqual("[EC, 00, 40, 00, 0F, 00, 00, 00, 58, 00, 00, 00, 0F, 00, 04, F0, 10, 00, 00, 00, 00, 00, 0A, F0, 08, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 0F, 00, 04, F0, 18, 00, 00, 00, 00, 00, 0A, F0, 08, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 11, F0, 00, 00, 00, 00, 5D, 00, 00, 00, EC, 00, 20, 00, 0F, 00, 04, F0, 18, 00, 00, 00, 00, 00, 0A, F0, 08, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 00, 11, F0, 00, 00, 00, 00, 5D, 00, 00, 00]",
                    HexDump.ToHex(data));
        }

    }
}