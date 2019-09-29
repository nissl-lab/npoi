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
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using NPOI.Util;

    /**
     * Tests for {@link RecordInputStream}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRecordInputStream
    {

        /**
         * Data inspired by attachment 22626 of bug 45866<br/>
         * A unicode string of 18 chars, with a continue record where the compression flag Changes
         */
        private static String HED_DUMP1 = ""
                + "1A 59 00 8A 9E 8A " // 3 uncompressed unicode chars
                + "3C 00 " // Continue sid
                + "10 00 " // rec size 16 (1+15)
                + "00"	// next chunk is compressed
                + "20 2D 20 4D 75 6C 74 69 6C 69 6E 67 75 61 6C " // 15 chars
        ;
        /**
         * same string re-arranged
         */
        private static String HED_DUMP2 = ""
            // 15 chars at end of current record
                + "4D 75 6C 74 69 6C 69 6E 67 75 61 6C 20 2D 20"
                + "3C 00 " // Continue sid
                + "07 00 " // rec size 7 (1+6)
                + "01"	// this bit uncompressed
                + "1A 59 00 8A 9E 8A " // 3 uncompressed unicode chars
        ;
        [Test]
        public void TestChangeOfCompressionFlag_bug25866()
        {
            byte[] changingFlagSimpleData = HexRead.ReadFromString(""
                    + "AA AA "  // fake SID
                    + "06 00 "  // first rec len 6
                    + HED_DUMP1
                    );
            RecordInputStream in1 = TestcaseRecordInputStream.Create(changingFlagSimpleData);
            String actual;
            try
            {
                actual = in1.ReadUnicodeLEString(18);
            }
            catch (ArgumentException e)
            {
                if ("compressByte in continue records must be 1 while Reading unicode LE string".Equals(e.Message))
                {
                    throw new AssertionException("Identified bug 45866");
                }

                throw e;
            }
            Assert.AreEqual("\u591A\u8A00\u8A9E - Multilingual", actual);
        }
        [Test]
        public void TestChangeFromUnCompressedToCompressed()
        {
            byte[] changingFlagSimpleData = HexRead.ReadFromString(""
                    + "AA AA "  // fake SID
                    + "0F 00 "  // first rec len 15
                    + HED_DUMP2
                    );
            RecordInputStream in1 = TestcaseRecordInputStream.Create(changingFlagSimpleData);
            String actual = in1.ReadCompressedUnicode(18);
            Assert.AreEqual("Multilingual - \u591A\u8A00\u8A9E", actual);
        }
        [Test]
        public void TestReadString()
        {
            byte[] changingFlagFullData = HexRead.ReadFromString(""
                    + "AA AA "  // fake SID
                    + "12 00 "  // first rec len 18 (15 + next 3 bytes)
                    + "12 00 "  // total chars 18
                    + "00 "	 // this bit compressed
                    + HED_DUMP2
                    );
            RecordInputStream in1 = TestcaseRecordInputStream.Create(changingFlagFullData);
            String actual = in1.ReadString();
            Assert.AreEqual("Multilingual - \u591A\u8A00\u8A9E", actual);
        }

        [Test]
        public void TestLeftoverDataException()
        {
            // just ensure that the exception is created correctly, even with unknown sids
            new LeftoverDataException(1, 200);
            new LeftoverDataException(0, 200);
            new LeftoverDataException(999999999, 200);
            new LeftoverDataException(HeaderRecord.sid, 200);
        }
    }

}