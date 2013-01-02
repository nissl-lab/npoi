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

namespace TestCases.HSSF.Record
{
    using NUnit.Framework;

    using NPOI.Util;
    using TestCases.HSSF.Record;
    using NPOI.HSSF.Record;

    /**
     * Tests for {@link ColumnInfoRecord}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestColumnInfoRecord
    {
        [Test]
        public void TestBasic() {
		byte[] data = HexRead.ReadFromString("7D 00 0C 00 14 00 9B 00 C7 19 0F 00 01 13 00 00");

		RecordInputStream in1 = TestcaseRecordInputStream.Create(data);
		ColumnInfoRecord cir = new ColumnInfoRecord(in1);
		Assert.AreEqual(0, in1.Remaining);

		Assert.AreEqual(20, cir.FirstColumn);
		Assert.AreEqual(155, cir.LastColumn);
		Assert.AreEqual(6599, cir.ColumnWidth);
		Assert.AreEqual(15, cir.XFIndex);
		Assert.AreEqual(true, cir.IsHidden);
		Assert.AreEqual(3, cir.OutlineLevel);
		Assert.AreEqual(true, cir.IsCollapsed);
		Assert.IsTrue(Arrays.Equals(data, cir.Serialize()));
	}

        /**
         * Some applications skip the last reserved field when writing {@link ColumnInfoRecord}s
         * The attached file was apparently Created by "SoftArtisans OfficeWriter for Excel".
         * Excel Reads that file OK and assumes zero for the value of the reserved field.
         */
        [Test]
        public void TestZeroResevedBytes_bug48332()
        {
            // Taken from bugzilla attachment 24661 (offset 0x1E73)
            byte[] inpData = HexRead.ReadFromString("7D 00 0A 00 00 00 00 00 D5 19 0F 00 02 00");
            byte[] outData = HexRead.ReadFromString("7D 00 0C 00 00 00 00 00 D5 19 0F 00 02 00 00 00");

            RecordInputStream in1 = TestcaseRecordInputStream.Create(inpData);
            ColumnInfoRecord cir;
            try
            {
                cir = new ColumnInfoRecord(in1);
            }
            catch (RuntimeException e)
            {
                if (e.Message.Equals("Unusual record size remaining=(0)"))
                {
                    throw new AssertionException("Identified bug 48332");
                }
                throw e;
            }
            Assert.AreEqual(0, in1.Remaining);
            Assert.IsTrue(Arrays.Equals(outData, cir.Serialize()));
        }

        /**
         * Some sample files have just one reserved byte (field 6):
         * OddStyleRecord.xls, NoGutsRecords.xls, WORKBOOK_in_capitals.xls
         * but this seems to cause no problem to Excel
         */
        [Test]
        public void TestOneReservedByte()
        {
            byte[] inpData = HexRead.ReadFromString("7D 00 0B 00 00 00 00 00 24 02 0F 00 00 00 01");
            byte[] outData = HexRead.ReadFromString("7D 00 0C 00 00 00 00 00 24 02 0F 00 00 00 01 00");
            RecordInputStream in1 = TestcaseRecordInputStream.Create(inpData);
            ColumnInfoRecord cir = new ColumnInfoRecord(in1);
            Assert.AreEqual(0, in1.Remaining);
            Assert.IsTrue(Arrays.Equals(outData, cir.Serialize()));
        }
    }

}