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
    using NPOI.HSSF.Record.CF;
    using NPOI.HSSF.Util;
    using NPOI.SS.Util;

    /**
     * Tests the serialization and deserialization of the TestCFHeaderRecord
     * class works correctly.  
     *
     * @author Dmitriy Kumshayev 
     */
    [TestFixture]
    public class TestCFHeaderRecord
    {
        [Test]
        public void TestCreateCFHeaderRecord()
        {
            CFHeaderRecord record = new CFHeaderRecord();
            CellRangeAddress[] ranges = {
			    new CellRangeAddress(0,0xFFFF,5,5),
			    new CellRangeAddress(0,0xFFFF,6,6),
			    new CellRangeAddress(0,1,0,1),
			    new CellRangeAddress(0,1,2,3),
			    new CellRangeAddress(2,3,0,1),
			    new CellRangeAddress(2,3,2,3),
		    };
            record.CellRanges = (ranges);
            ranges = record.CellRanges;
            Assert.AreEqual(6, ranges.Length);
            CellRangeAddress enclosingCellRange = record.EnclosingCellRange;
            Assert.AreEqual(0, enclosingCellRange.FirstRow);
            Assert.AreEqual(65535, enclosingCellRange.LastRow);
            Assert.AreEqual(0, enclosingCellRange.FirstColumn);
            Assert.AreEqual(6, enclosingCellRange.LastColumn);

            Assert.AreEqual(false, record.NeedRecalculation);
            Assert.AreEqual(0, record.ID);

            record.NeedRecalculation = (true);
            Assert.AreEqual(true, record.NeedRecalculation);
            Assert.AreEqual(0, record.ID);

            record.ID = (7);
            record.NeedRecalculation = (false);
            Assert.AreEqual(false, record.NeedRecalculation);
            Assert.AreEqual(7, record.ID);

        }

        [Test]
        public void TestCreateCFHeader12Record()
        {
            CFHeader12Record record = new CFHeader12Record();
            CellRangeAddress[] ranges = {
            new CellRangeAddress(0,0xFFFF,5,5),
            new CellRangeAddress(0,0xFFFF,6,6),
            new CellRangeAddress(0,1,0,1),
            new CellRangeAddress(0,1,2,3),
            new CellRangeAddress(2,3,0,1),
            new CellRangeAddress(2,3,2,3),
        };
            record.CellRanges = (ranges);
            ranges = record.CellRanges;
            Assert.AreEqual(6, ranges.Length);
            CellRangeAddress enclosingCellRange = record.EnclosingCellRange;
            Assert.AreEqual(0, enclosingCellRange.FirstRow);
            Assert.AreEqual(65535, enclosingCellRange.LastRow);
            Assert.AreEqual(0, enclosingCellRange.FirstColumn);
            Assert.AreEqual(6, enclosingCellRange.LastColumn);
            Assert.AreEqual(false, record.NeedRecalculation);
            Assert.AreEqual(0, record.ID);

            record.NeedRecalculation = (true);
            Assert.AreEqual(true, record.NeedRecalculation);
            Assert.AreEqual(0, record.ID);

            record.ID = (7);
            record.NeedRecalculation = (false);
            Assert.AreEqual(false, record.NeedRecalculation);
            Assert.AreEqual(7, record.ID);
        }


        [Test]
        public void TestSerialization()
        {
            byte[] recordData = 
		    {
			    (byte)0x03, (byte)0x00,
			    (byte)0x01,	(byte)0x00,
    			
			    (byte)0x00,	(byte)0x00,
			    (byte)0x03,	(byte)0x00,
			    (byte)0x00,	(byte)0x00,
			    (byte)0x03,	(byte)0x00,
    			
			    (byte)0x04,	(byte)0x00, // nRegions
    			
			    (byte)0x00,	(byte)0x00,
			    (byte)0x01,	(byte)0x00,
			    (byte)0x00,	(byte)0x00,
			    (byte)0x01,	(byte)0x00,
    			
			    (byte)0x00,	(byte)0x00,
			    (byte)0x01,	(byte)0x00,
			    (byte)0x02,	(byte)0x00,
			    (byte)0x03,	(byte)0x00,
    			
			    (byte)0x02,	(byte)0x00,
			    (byte)0x03,	(byte)0x00,
			    (byte)0x00,	(byte)0x00,
			    (byte)0x01,	(byte)0x00,
    			
			    (byte)0x02,	(byte)0x00,
			    (byte)0x03,	(byte)0x00,
			    (byte)0x02,	(byte)0x00,
			    (byte)0x03,	(byte)0x00,
		    };

            CFHeaderRecord record = new CFHeaderRecord(TestcaseRecordInputStream.Create(CFHeaderRecord.sid, recordData));

            Assert.AreEqual(3, record.NumberOfConditionalFormats, "#CFRULES");
            Assert.IsTrue(record.NeedRecalculation);
            Confirm(record.EnclosingCellRange, 0, 3, 0, 3);
            CellRangeAddress[] ranges = record.CellRanges;
            Assert.AreEqual(4, ranges.Length);
            Confirm(ranges[0], 0, 1, 0, 1);
            Confirm(ranges[1], 0, 1, 2, 3);
            Confirm(ranges[2], 2, 3, 0, 1);
            Confirm(ranges[3], 2, 3, 2, 3);
            Assert.AreEqual(recordData.Length + 4, record.RecordSize);

            byte[] output = record.Serialize();

            Assert.AreEqual(recordData.Length + 4, output.Length, "Output size"); //includes sid+recordlength

            for (int i = 0; i < recordData.Length; i++)
            {
                Assert.AreEqual(recordData[i], output[i + 4], "CFHeaderRecord doesn't match");
            }
        }
        [Test]
        public void TestExtremeRows()
        {
            byte[] recordData = {
			    (byte)0x13, (byte)0x00, // nFormats
			    (byte)0x00,	(byte)0x00,
    			
			    (byte)0x00,	(byte)0x00,
			    (byte)0xFF,	(byte)0xFF,
			    (byte)0x00,	(byte)0x00,
			    (byte)0xFF,	(byte)0x00,
    			
			    (byte)0x03,	(byte)0x00, // nRegions
    			
			    (byte)0x40,	(byte)0x9C,
			    (byte)0x50,	(byte)0xC3,
			    (byte)0x02,	(byte)0x00,
			    (byte)0x02,	(byte)0x00,
    			
			    (byte)0x00,	(byte)0x00,
			    (byte)0xFF,	(byte)0xFF,
			    (byte)0x05,	(byte)0x00,
			    (byte)0x05,	(byte)0x00,
    			
			    (byte)0x07,	(byte)0x00,
			    (byte)0x07,	(byte)0x00,
			    (byte)0x00,	(byte)0x00,
			    (byte)0xFF,	(byte)0x00,
		    };

            CFHeaderRecord record;
            try
            {
                record = new CFHeaderRecord(TestcaseRecordInputStream.Create(CFHeaderRecord.sid, recordData));
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("invalid cell range (-25536, 2, -15536, 2)"))
                {
                    throw new AssertionException("Identified bug 44739b");
                }
                throw e;
            }

            Assert.AreEqual(19, record.NumberOfConditionalFormats, "#CFRULES");
            Assert.IsFalse(record.NeedRecalculation);
            Confirm(record.EnclosingCellRange, 0, 65535, 0, 255);
            CellRangeAddress[] ranges = record.CellRanges;
            Assert.AreEqual(3, ranges.Length);
            Confirm(ranges[0], 40000, 50000, 2, 2);
            Confirm(ranges[1], 0, 65535, 5, 5);
            Confirm(ranges[2], 7, 7, 0, 255);

            byte[] output = record.Serialize();

            Assert.AreEqual(recordData.Length + 4, output.Length, "Output size"); //includes sid+recordlength

            for (int i = 0; i < recordData.Length; i++)
            {
                Assert.AreEqual(recordData[i], output[i + 4], "CFHeaderRecord doesn't match");
            }
        }


        private static void Confirm(CellRangeAddress cr, int expFirstRow, int expLastRow, int expFirstCol, int expLastColumn)
        {
            Assert.AreEqual(expFirstRow, cr.FirstRow, "first row");
            Assert.AreEqual(expLastRow, cr.LastRow, "last row");
            Assert.AreEqual(expFirstCol, cr.FirstColumn, "first column");
            Assert.AreEqual(expLastColumn, cr.LastColumn, "last column");
        }

    }
}
