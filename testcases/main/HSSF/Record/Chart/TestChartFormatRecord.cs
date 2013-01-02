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

namespace TestCases.HSSF.Record.Chart
{
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.Util;
    using TestCases.HSSF.Record;

    /**
     * Tests for {@link ChartFormatRecord} Test data taken directly from a real
     * Excel file.
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestChartFormatRecord
    {
        /**
         * This rather uninteresting data came from attachment 23347 of bug 46693 at
         * offsets 0x6BB2 and 0x7BAF
         */
        private static byte[] data = HexRead.ReadFromString(
                "14 10 14 00 " // BIFF header
                + "00 00 00 00 00 00 00 00 "
                + "00 00 00 00 00 00 00 00 "
                + "00 00 00 00");

        /**
         * The correct size of a {@link ChartFormatRecord} is 20 bytes (not including header).
         */
        [Test]
        public void TestLoad()
        {
            RecordInputStream in1 = TestcaseRecordInputStream.Create(data);
            ChartFormatRecord record = new ChartFormatRecord(in1);
            if (in1.Remaining == 2)
            {
                throw new AssertionException("Identified bug 44693d");
            }
            Assert.AreEqual(0, in1.Remaining);
            Assert.AreEqual(24, record.RecordSize);

            byte[] data2 = record.Serialize();
            Assert.IsTrue(Arrays.Equals(data, data2));
        }
    }

}