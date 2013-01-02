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
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRecalcIdRecord
    {
        [Test]
        public void TestBasicDeSerializeReserialize()
        {

            byte[] data = HexRead.ReadFromString(
                    "C1 01" +  // rt
                    "00 00" +  // reserved
                    "1D EB 01 00"); // engine id

            RecalcIdRecord r = Create(data);
            TestcaseRecordInputStream.ConfirmRecordEncoding(RecalcIdRecord.sid, data, r.Serialize());
        }
        [Test]
        public void TestBadFirstField_bug48096()
        {
            /**
             * Data taken from the sample file referenced in Bugzilla 48096, file offset 0x0D45.
             * The apparent problem is that the first data short field has been written with the
             * wrong <i>endianness</n>.  Excel seems to ignore whatever value is present in this
             * field, and always reWrites it as (C1, 01). POI now does the same.
             */
            byte[] badData = HexRead.ReadFromString("C1 01 08 00 01 C1 00 00 00 01 69 61");
            byte[] goodData = HexRead.ReadFromString("C1 01 08 00 C1 01 00 00 00 01 69 61");

            RecordInputStream in1 = TestcaseRecordInputStream.Create(badData);
            RecalcIdRecord r;
            try
            {
                r = new RecalcIdRecord(in1);
            }
            catch (RecordFormatException e)
            {
                if (e.Message.Equals("expected 449 but got 49409"))
                {
                    throw new AssertionException("Identified bug 48096");
                }
                throw e;
            }
            Assert.AreEqual(0, in1.Remaining);
            Assert.IsTrue(Arrays.Equals(r.Serialize(), goodData));
        }
        private static RecalcIdRecord Create(byte[] data)
        {
            RecordInputStream in1 = TestcaseRecordInputStream.Create(RecalcIdRecord.sid, data);
            RecalcIdRecord result = new RecalcIdRecord(in1);
            Assert.AreEqual(0, in1.Remaining);
            return result;
        }
    }

}