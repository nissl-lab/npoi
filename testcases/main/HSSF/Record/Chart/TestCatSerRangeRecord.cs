
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



namespace TestCases.HSSF.Record.Chart
{
    using System;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Chart;
    using NUnit.Framework;using NUnit.Framework.Legacy;

    /**
     * Tests the serialization and deserialization of the CategorySeriesAxisRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestCatSerRangeRecord
    {
        byte[] data = new byte[] {
        (byte)0x01,(byte)0x00,    // crossing point
        (byte)0x01,(byte)0x00,    // label frequency
        (byte)0x01,(byte)0x00,    // tick mark frequency
        (byte)0x01,(byte)0x00     // options
    };

        public TestCatSerRangeRecord()
        {

        }
        [Test]
        public void TestLoad()
        {

            CatSerRangeRecord record = new CatSerRangeRecord(TestcaseRecordInputStream.Create((short)0x1020, data));
            ClassicAssert.AreEqual(1, record.CrossPoint);
            ClassicAssert.AreEqual(1, record.LabelInterval);
            ClassicAssert.AreEqual(1, record.MarkInterval);
            ClassicAssert.AreEqual(1, record.Options);
            ClassicAssert.AreEqual(true, record.IsBetween);
            ClassicAssert.AreEqual(false, record.IsMaxCross);
            ClassicAssert.AreEqual(false, record.IsReverse);


            ClassicAssert.AreEqual(4 + 8, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            CatSerRangeRecord record = new CatSerRangeRecord();
            record.CrossPoint = ((short)1);
            record.LabelInterval = ((short)1);
            record.MarkInterval = ((short)1);
            record.IsBetween = (true);
            record.IsMaxCross = (false);
            record.IsReverse = (false);


            byte[] recordBytes = record.Serialize();
            ClassicAssert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                ClassicAssert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}