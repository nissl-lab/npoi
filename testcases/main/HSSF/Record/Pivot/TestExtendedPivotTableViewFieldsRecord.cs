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

namespace TestCases.HSSF.Record.Pivot
{

    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.PivotTable;
    using NPOI.Util;
    using TestCases.HSSF.Record;

    /**
     * Tests for {@link ExtendedPivotTableViewFieldsRecord}
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestExtendedPivotTableViewFieldsRecord
    {

        [Test]
        public void TestSubNameNotPresent_bug46693()
        {
            // This data came from attachment 23347 of bug 46693 at offset 0xAA43
            byte[] data = HexRead.ReadFromString(
                    "00 01 14 00" + // BIFF header
                    "1E 14 00 0A FF FF FF FF 00 00 FF FF 00 00 00 00 00 00 00 00");
            RecordInputStream in1 = TestcaseRecordInputStream.Create(data);
            ExtendedPivotTableViewFieldsRecord rec;
            try
            {
                rec = new ExtendedPivotTableViewFieldsRecord(in1);
            }
            catch (RecordFormatException e)
            {
                if (e.Message.Equals("Expected to find a ContinueRecord in order to read remaining 65535 of 65535 chars"))
                {
                    throw new AssertionException("Identified bug 46693a");
                }
                throw e;
            }

            Assert.AreEqual(data.Length, rec.RecordSize);
        }

        [Test]
        public void TestOlderFormat_bug46918()
        {
            // There are 10 SXVDEX records in the file (not uploaded) that originated bugzilla 46918
            // They all had the following hex encoding:
            byte[] data = HexRead.ReadFromString("00 01 0A 00 1E 14 00 0A FF FF FF FF 00 00");

            RecordInputStream in1 = TestcaseRecordInputStream.Create(data);
            ExtendedPivotTableViewFieldsRecord rec;
            try
            {
                rec = new ExtendedPivotTableViewFieldsRecord(in1);
            }
            catch (RecordFormatException e)
            {
                if (e.Message.Equals("Not enough data (0) to read requested (2) bytes"))
                {
                    throw new AssertionException("Identified bug 46918");
                }
                throw e;
            }

            byte[] expReserData = HexRead.ReadFromString("1E 14 00 0A FF FF FF FF 00 00" +
                    "FF FF 00 00 00 00 00 00 00 00");

            TestcaseRecordInputStream.ConfirmRecordEncoding(ExtendedPivotTableViewFieldsRecord.sid, expReserData, rec.Serialize());
        }
    }

}