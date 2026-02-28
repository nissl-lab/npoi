
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
     * Tests the serialization and deserialization of the LegendRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     *

     * @author Andrew C. Oliver (acoliver at apache.org)
     */
    [TestFixture]
    public class TestLegendRecord
    {
        byte[] data = new byte[] {
	(byte)0x76,(byte)0x0E,(byte)0x00,(byte)0x00,(byte)0x86,(byte)0x07,(byte)0x00,(byte)0x00,(byte)0x19,(byte)0x01,(byte)0x00,(byte)0x00,(byte)0x8B,(byte)0x00,(byte)0x00,(byte)0x00,(byte)0x03,(byte)0x01,(byte)0x1F,(byte)0x00
    };

        public TestLegendRecord()
        {

        }
        [Test]
        public void TestLoad()
        {
            LegendRecord record = new LegendRecord(TestcaseRecordInputStream.Create((short)0x1015, data));


            ClassicAssert.AreEqual((int)0xe76, record.XAxisUpperLeft);

            ClassicAssert.AreEqual((int)0x786, record.YAxisUpperLeft);

            ClassicAssert.AreEqual((int)0x119, record.XSize);

            ClassicAssert.AreEqual((int)0x8b, record.YSize);

            ClassicAssert.AreEqual((byte)0x3, record.Type);

            ClassicAssert.AreEqual((byte)0x1, record.Spacing);

            ClassicAssert.AreEqual((short)0x1f, record.Options);
            ClassicAssert.AreEqual(true, record.IsAutoPosition);
            ClassicAssert.AreEqual(true, record.IsAutoSeries);
            ClassicAssert.AreEqual(true, record.IsAutoXPositioning);
            ClassicAssert.AreEqual(true, record.IsAutoYPositioning);
            ClassicAssert.AreEqual(true, record.IsVertical);
            ClassicAssert.AreEqual(false, record.IsDataTable);


            ClassicAssert.AreEqual(24, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            LegendRecord record = new LegendRecord();



            record.XAxisUpperLeft=((int)0xe76);

            record.YAxisUpperLeft=((int)0x786);

            record.XSize = ((int)0x119);

            record.YSize = ((int)0x8b);

            record.Type = ((byte)0x3);

            record.Spacing = ((byte)0x1);

            record.Options = ((short)0x1f);
            record.IsAutoPosition = (true);
            record.IsAutoSeries = (true);
            record.IsAutoXPositioning = (true);
            record.IsAutoYPositioning = (true);
            record.IsVertical = (true);
            record.IsDataTable = (false);


            byte[] recordBytes = record.Serialize();
            ClassicAssert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                ClassicAssert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }
}