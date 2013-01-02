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
    using NPOI.Util;
    using TestCases.HSSF.Record;
    using NUnit.Framework;
    using NPOI.HSSF.Record;


    /**
     * Tests for {@link WriteAccessRecord}
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestWriteAccessRecord
    {

        private static String HEX_SIXTYFOUR_SPACES = ""
            + "20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 "
            + "20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 "
            + "20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 "
            + "20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20";

        [Test]
        public void TestMissingStringHeader_bug47001a()
        {
            /*
             * Data taken from offset 0x0224 in
             * attachment 23468 from bugzilla 47001
             */
            byte[] data = HexRead.ReadFromString(""
                    + "5C 00 70 00 "
                    + "4A 61 76 61 20 45 78 63 65 6C 20 41 50 49 20 76 "
                    + "32 2E 36 2E 34"
                    + "20 20 20 20 20 20 20 20 20 20 20 "
                    + "20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 "
                    + HEX_SIXTYFOUR_SPACES);

            RecordInputStream in1 = TestcaseRecordInputStream.Create(data);

            WriteAccessRecord rec;
            try
            {
                rec = new WriteAccessRecord(in1);
            }
            catch (RecordFormatException e)
            {
                if (e.Message.Equals("Not enough data (0) to read requested (1) bytes"))
                {
                    throw new AssertionException("Identified bug 47001a");
                }
                throw e;
            }
            Assert.AreEqual("Java Excel API v2.6.4", rec.Username);


            byte[] expectedEncoding = HexRead.ReadFromString(""
                    + "15 00 00 4A 61 76 61 20 45 78 63 65 6C 20 41 50 "
                    + "49 20 76 32 2E 36 2E 34"
                    + "20 20 20 20 20 20 20 20 "
                    + "20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 "
                    + HEX_SIXTYFOUR_SPACES);

            TestcaseRecordInputStream.ConfirmRecordEncoding(WriteAccessRecord.sid, expectedEncoding, rec.Serialize());
        }
        [Test]
        public void TestshortRecordWrittenByMSAccess()
        {
            /*
             * Data taken from two example files
             * ex42564-21435.xls
             * bug_42794.xls (from bug 42794 attachment 20429)
             * In both cases, this data is found at offset 0x0C1C.
             */
            byte[] data = HexRead.ReadFromString(""
                    + "5C 00 39 00 "
                    + "36 00 00 41 20 73 61 74 69 73 66 69 65 64 20 4D "
                    + "69 63 72 6F 73 6F 66 74 20 4F 66 66 69 63 65 39 "
                    + "20 55 73 65 72"
                    + "20 20 20 20 20 20 20 20 20 20 20 "
                    + "20 20 20 20 20 20 20 20 20");

            RecordInputStream in1 = TestcaseRecordInputStream.Create(data);
            WriteAccessRecord rec = new WriteAccessRecord(in1);
            Assert.AreEqual("A satisfied Microsoft Office9 User", rec.Username);
            byte[] expectedEncoding = HexRead.ReadFromString(""
                    + "22 00 00 41 20 73 61 74 69 73 66 69 65 64 20 4D "
                    + "69 63 72 6F 73 6F 66 74 20 4F 66 66 69 63 65 39 "
                    + "20 55 73 65 72"
                    + "20 20 20 20 20 20 20 20 20 20 20 "
                    + HEX_SIXTYFOUR_SPACES);

            TestcaseRecordInputStream.ConfirmRecordEncoding(WriteAccessRecord.sid, expectedEncoding, rec.Serialize());
        }
    }

}
