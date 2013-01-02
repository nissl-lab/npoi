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
    using TestCases.HSSF.Record;
    using NPOI.Util;
    using NPOI.HSSF.Record;
    using NUnit.Framework;

    /**
     * Tests for {@link BoolErrRecord}
     */
    [TestFixture]
    public class TestBoolErrRecord
    {
        [Test]
        public void TestError()
        {
            byte[] data = HexRead.ReadFromString(
                    "00 00 00 00 0F 00 " + // row, col, xfIndex
                    "07 01 " // #DIV/0!, isError
                    );

            RecordInputStream in1 = TestcaseRecordInputStream.Create(BoolErrRecord.sid, data);
            BoolErrRecord ber = new BoolErrRecord(in1);
            Assert.IsTrue(ber.IsError);
            Assert.AreEqual(7, ber.ErrorValue);

            TestcaseRecordInputStream.ConfirmRecordEncoding(BoolErrRecord.sid, data, ber.Serialize());
        }

        /**
         * Bugzilla 47479 was due to an apparent error in OOO which (as of version 3.0.1) 
         * Writes the <i>value</i> field of BOOLERR records as 2 bytes instead of 1.<br/>
         * Coincidentally, the extra byte written is zero, which is exactly the value 
         * required by the <i>isError</i> field.  This probably why Excel seems to have
         * no problem.  OOO does not have the same bug for error values (which wouldn't
         * work by the same coincidence). 
         */
        [Test]
        public void TestOooBadFormat_bug47479()
        {
            byte[] data = HexRead.ReadFromString(
                    "05 02 09 00 " + // sid, size
                    "00 00 00 00 0F 00 " + // row, col, xfIndex
                    "01 00 00 " // extra 00 byte here
                    );

            RecordInputStream in1 = TestcaseRecordInputStream.Create(data);
            BoolErrRecord ber = new BoolErrRecord(in1);
            bool hasMore;
            try
            {
                hasMore = in1.HasNextRecord;
            }
            catch (LeftoverDataException e)
            {
                if ("Initialisation of record 0x205 left 1 bytes remaining still to be Read.".Equals(e.Message))
                {
                    throw new AssertionException("Identified bug 47479");
                }
                throw e;
            }
            Assert.IsFalse(hasMore);
            Assert.IsTrue(ber.IsBoolean);
            Assert.AreEqual(true, ber.BooleanValue);

            // Check that the record re-Serializes correctly
            byte[] outData = ber.Serialize();
            byte[] expData = HexRead.ReadFromString(
                    "05 02 08 00 " +
                    "00 00 00 00 0F 00 " +
                    "01 00 " // normal number of data bytes
                    );
            Assert.IsTrue(Arrays.Equals(expData, outData));
        }
    }
}


