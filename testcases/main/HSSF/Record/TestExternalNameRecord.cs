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
    using NPOI.Util;
    /**
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestExternalNameRecord
    {

        private static byte[] dataFDS = {
            0, 0, 0, 0, 0, 0, 3, 0, 70, 68, 83, 0, 0,
        };

        // data taken from bugzilla 44774 att 21790
        private static byte[] dataAutoDocName = {
            unchecked((byte)-22), 127, 0, 0, 0, 0, 29, 0, 39, 49, 57, 49, 50, 49, 57, 65, 87, 52, 32, 67, 111, 114,
            112, 44, 91, 87, 79, 82, 75, 79, 85, 84, 95, 80, 88, 93, 39,
        };

        // data taken from bugzilla 44774 att 21790
        private static byte[] dataPlainName = {
            0, 0, 0, 0, 0, 0, 9, 0, 82, 97, 116, 101, 95, 68, 97, 116, 101, 9, 0, 58, 0, 0, 0, 0, 4, 0, 8, 0
            // TODO - the last 2 bytes of formula data (8,0) seem weird.  They encode to ConcatPtg, UnknownPtg
            // UnknownPtg is1 otherwise not created by any other test cases
        };

        private static ExternalNameRecord CreateSimpleENR(byte[] data)
        {
            return new ExternalNameRecord(TestcaseRecordInputStream.Create(0x0023, data));
        }
        [Test]
        public void TestBasicDeserializeReserialize()
        {

            ExternalNameRecord enr = CreateSimpleENR(dataFDS);
            Assert.AreEqual("FDS", enr.Text);

            try
            {
                TestcaseRecordInputStream.ConfirmRecordEncoding(0x0023, dataFDS, enr.Serialize());
            }
            catch (IndexOutOfRangeException e)
            {
                if (e.Message.Equals("15"))
                {
                    throw new AssertionException("Identified bug 44695");
                }
            }
        }
        [Test]
        public void TestBasicSize()
        {
            ExternalNameRecord enr = CreateSimpleENR(dataFDS);
            if (enr.RecordSize == 13)
            {
                throw new AssertionException("Identified bug 44695");
            }
            Assert.AreEqual(17, enr.RecordSize);

            Assert.IsNotNull(enr.Serialize());
        }
        [Test]
        public void TestAutoStdDocName()
        {

            ExternalNameRecord enr;
            try
            {
                enr = CreateSimpleENR(dataAutoDocName);
            }
            catch (IndexOutOfRangeException e)
            {
                if (e.Message == null)
                {
                    throw new AssertionException("Identified bug XXXX");
                }
                throw e;
            }
            Assert.AreEqual("'191219AW4 Corp,[WORKOUT_PX]'", enr.Text);
            Assert.IsTrue(enr.IsAutomaticLink);
            Assert.IsFalse(enr.IsBuiltInName);
            Assert.IsFalse(enr.IsIconifiedPictureLink);

            Assert.IsFalse(enr.IsOLELink);
            Assert.IsFalse(enr.IsPicureLink);
            Assert.IsTrue(enr.IsStdDocumentNameIdentifier);


            TestcaseRecordInputStream.ConfirmRecordEncoding(0x0023, dataAutoDocName, enr.Serialize());
        }
        [Test]
        public void TestPlainName()
        {

            ExternalNameRecord enr = CreateSimpleENR(dataPlainName);
            Assert.AreEqual("Rate_Date", enr.Text);
            Assert.IsFalse(enr.IsAutomaticLink);
            Assert.IsFalse(enr.IsBuiltInName);
            Assert.IsFalse(enr.IsIconifiedPictureLink);
            Assert.IsFalse(enr.IsOLELink);
            Assert.IsFalse(enr.IsPicureLink);
            Assert.IsFalse(enr.IsStdDocumentNameIdentifier);

            TestcaseRecordInputStream.ConfirmRecordEncoding(0x0023, dataPlainName, enr.Serialize());
        }
        [Test]
        public void TestDDELink_bug47229()
        {
            /**
             * Hex dump read directly from text of bugzilla 47229
             */
            byte[] dataDDE = NPOI.Util.HexRead.ReadFromString(
                   "E2 7F 00 00 00 00 " +
                   "37 00 " + // text len
                // 010672AT0 MUNI,[RTG_MOODY_UNDERLYING,RTG_SP_UNDERLYING]
                   "30 31 30 36 37 32 41 54 30 20 4D 55 4E 49 2C " +
                   "5B 52 54 47 5F 4D 4F 4F 44 59 5F 55 4E 44 45 52 4C 59 49 4E 47 2C " +
                   "52 54 47 5F 53 50 5F 55 4E 44 45 52 4C 59 49 4E 47 5D " +
                // constant array { { "#N/A N.A.", "#N/A N.A.", }, }
                   " 01 00 00 " +
                   "02 09 00 00 23 4E 2F 41 20 4E 2E 41 2E " +
                   "02 09 00 00 23 4E 2F 41 20 4E 2E 41 2E");
            ExternalNameRecord enr;
            try
            {
                enr = CreateSimpleENR(dataDDE);
            }
            catch (RecordFormatException e)
            {
                // actual msg reported in bugzilla 47229 is different
                // because that seems to be using a version from before svn r646666
                if (e.Message.StartsWith("Some unread data (is formula present?)"))
                {
                    throw new AssertionException("Identified bug 47229 - failed to read ENR with OLE/DDE result data");
                }
                throw e;
            }
            Assert.AreEqual("010672AT0 MUNI,[RTG_MOODY_UNDERLYING,RTG_SP_UNDERLYING]", enr.Text);

            TestcaseRecordInputStream.ConfirmRecordEncoding(0x0023, dataDDE, enr.Serialize());
        }
        [Test]
        public void TestUnicodeName_bug47384()
        {
            // data taken from bugzilla 47384 att 23830 at offset 0x13A0
            byte[] dataUN = NPOI.Util.HexRead.ReadFromString(
                    "23 00 22 00" +
                    "00 00 00 00 00 00 " +
                    "0C 01 " +
                    "59 01 61 00 7A 00 65 00 6E 00 ED 00 5F 00 42 00 69 00 6C 00 6C 00 61 00" +
                    "00 00");

            RecordInputStream in1 = TestcaseRecordInputStream.Create(dataUN);
            ExternalNameRecord enr;
            try
            {
                enr = new ExternalNameRecord(in1);
            }
            catch (RecordFormatException e)
            {
                if (e.Message.StartsWith("Expected to find a ContinueRecord in order to read remaining 242 of 268 chars"))
                {
                    throw new AssertionException("Identified bug 47384 - failed to read ENR with unicode name");
                }
                throw e;
            }
            Assert.AreEqual("\u0159azen\u00ED_Billa", enr.Text);
        }

        [Test]
        public void TestNPEWithFileFrom49219()
        {
            // the file at test-data/spreadsheet/49219.xls has ExternalNameRecords without actual data, 
            // we did handle this during Reading, but failed during serializing this out1, ensure it works now
            byte[] data = new byte[] {
        		2, 127, 0, 0, 0, 0, 
        		9, 0, 82, 97, 116, 101, 95, 68, 97, 116, 101};

            ExternalNameRecord enr = CreateSimpleENR(data);

            byte[] ser = enr.Serialize();
            Assert.AreEqual("[23, 00, 11, 00, 02, 7F, 00, 00, 00, 00, 09, 00, 52, 61, 74, 65, 5F, 44, 61, 74, 65]",
                    HexDump.ToHex(ser));
        }

    }
}