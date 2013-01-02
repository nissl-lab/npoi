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
    using System;

    using NPOI.HSSF.Util;

    using NUnit.Framework;
    using NPOI.SS.Util;
    using TestCases.HSSF.Record;
    using NPOI.HSSF.Record;

    /**
     * Tests the serialization and deserialization of the TableRecord
     * class works correctly.  Test data taken directly from a real
     * Excel file.
     */
    [TestFixture]
    public class TestTableRecord
    {
        byte[] header = new byte[] {
			0x36, 02, 0x10, 00, // sid=x236, 16 bytes long
	};
        byte[] data = new byte[] {
			03, 00,  // from row 3 
			8, 00,   // to row 8
			04,      // from col 4
			06,      // to col 6
			00, 00,  // no flags Set
			04, 00,  // row inp row 4
			01, 00,  // col inp row 1
			0x76, 0x40, // row inp col 0x4076 (!)
			00, 00   // col inp col 0
	};
        [Test]
        public void TestLoad()
        {

            TableRecord record = new TableRecord(TestcaseRecordInputStream.Create(0x236, data));

            CellRangeAddress8Bit range = record.Range;
            Assert.AreEqual(3, range.FirstRow);
            Assert.AreEqual(8, range.LastRow);
            Assert.AreEqual(4, range.FirstColumn);
            Assert.AreEqual(6, range.LastColumn);
            Assert.AreEqual(0, record.Flags);
            Assert.AreEqual(4, record.RowInputRow);
            Assert.AreEqual(1, record.ColInputRow);
            Assert.AreEqual(0x4076, record.RowInputCol);
            Assert.AreEqual(0, record.ColInputCol);

            Assert.AreEqual(16 + 4, record.RecordSize);
        }
        [Test]
        public void TestStore()
        {
            //    	Offset 0x3bd9 (15321)
            //    	recordid = 0x236, size = 16
            //    	[TABLE]
            //    	    .row from      = 3
            //    	    .row to        = 8
            //    	    .column from   = 4
            //    	    .column to     = 6
            //    	    .flags         = 0
            //    	        .always calc     =false
            //    	    .reserved      = 0
            //    	    .row input row = 4
            //    	    .col input row = 1
            //    	    .row input col = 4076
            //    	    .col input col = 0
            //    	[/TABLE]

            CellRangeAddress8Bit crab = new CellRangeAddress8Bit(3, 8, 4, 6);
            TableRecord record = new TableRecord(crab);
            record.Flags = (/*setter*/(byte)0);
            record.RowInputRow = (/*setter*/4);
            record.ColInputRow = (/*setter*/1);
            record.RowInputCol = (/*setter*/0x4076);
            record.ColInputCol = (/*setter*/0);

            byte[] recordBytes = record.Serialize();
            Assert.AreEqual(recordBytes.Length - 4, data.Length);
            for (int i = 0; i < data.Length; i++)
                Assert.AreEqual(data[i], recordBytes[i + 4], "At offset " + i);
        }
    }

}