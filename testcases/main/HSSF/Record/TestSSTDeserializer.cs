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
    using System.IO;
    using NUnit.Framework;
    using NPOI.HSSF.Record;
    using TestCases.HSSF;
    using NPOI.Util;
    using NPOI.Util.Collections;

    /**
     * Exercise the SSTDeserializer class.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestSSTDeserializer
    {
        private static int FAKE_SID = -5555;

        private static byte[] Concat(byte[] a, byte[] b)
        {
            byte[] result = new byte[a.Length + b.Length];
            System.Array.Copy(a, 0, result, 0, a.Length);
            System.Array.Copy(b, 0, result, a.Length, b.Length);
            return result;
        }

        private static byte[] ReadSampleHexData(String sampleFileName, String sectionName,int recSid)
        {
            Stream is1 = HSSFTestDataSamples.OpenSampleFileStream(sampleFileName);
            byte[] data;
            try {
			    data = HexRead.ReadData(is1, sectionName);
            }
            catch (IOException)
            {
                throw;
            }
            return TestcaseRecordInputStream.MergeDataAndSid(recSid, data.Length, data);
        }
        [Test]
        public void TestSpanRichTextToPlainText()
        {
            byte[] header = ReadSampleHexData("richtextdata.txt", "header", FAKE_SID);
            byte[] continueBytes = ReadSampleHexData("richtextdata.txt", "continue1", ContinueRecord.sid);
            RecordInputStream in1 = TestcaseRecordInputStream.Create(Concat(header, continueBytes));


            IntMapper<UnicodeString> strings = new IntMapper<UnicodeString>();
            SSTDeserializer deserializer = new SSTDeserializer(strings);
            deserializer.ManufactureStrings(1, in1);

            Assert.AreEqual("At a dinner party orAt At At ", strings[0] + "");
        }
        [Test]
        public void TestContinuationWithNoOverlap()
        {
            byte[] header = ReadSampleHexData("evencontinuation.txt", "header", FAKE_SID);
            byte[] continueBytes = ReadSampleHexData("evencontinuation.txt", "continue1", ContinueRecord.sid);
            RecordInputStream in1 = TestcaseRecordInputStream.Create(Concat(header, continueBytes));

            IntMapper<UnicodeString> strings = new IntMapper<UnicodeString>();
            SSTDeserializer deserializer = new SSTDeserializer(strings);
            deserializer.ManufactureStrings(2, in1);

            Assert.AreEqual("At a dinner party or", strings[0] + "");
            Assert.AreEqual("At a dinner party", strings[1] + "");
        }

        /**
         * Strings can actually span across more than one continuation.
         */
        [Test]
        public void TestStringAcross2Continuations()
        {
            byte[] header = ReadSampleHexData("stringacross2continuations.txt", "header", FAKE_SID);
            byte[] continue1 = ReadSampleHexData("stringacross2continuations.txt", "continue1", ContinueRecord.sid);
            byte[] continue2 = ReadSampleHexData("stringacross2continuations.txt", "continue2", ContinueRecord.sid);

            byte[] bytes = Concat(header, continue1);
            bytes = Concat(bytes, continue2);
            RecordInputStream in1 = TestcaseRecordInputStream.Create(bytes);

            IntMapper<UnicodeString> strings = new IntMapper<UnicodeString>();
            SSTDeserializer deserializer = new SSTDeserializer(strings);
            deserializer.ManufactureStrings(2, in1);

            Assert.AreEqual("At a dinner party or", strings[0] + "");
            Assert.AreEqual("At a dinner partyAt a dinner party", strings[1] + "");
        }
        [Test]
        public void TestExtendedStrings()
        {
            byte[] header = ReadSampleHexData("extendedtextstrings.txt", "rich-header", FAKE_SID);
            byte[] continueBytes = ReadSampleHexData("extendedtextstrings.txt", "rich-continue1", ContinueRecord.sid);
            RecordInputStream in1 = TestcaseRecordInputStream.Create(Concat(header, continueBytes));

            IntMapper<UnicodeString> strings = new IntMapper<UnicodeString>();
            SSTDeserializer deserializer = new SSTDeserializer(strings);
            deserializer.ManufactureStrings(1, in1);

            Assert.AreEqual("At a dinner party orAt At At ", strings[0].ToString());


            header = ReadSampleHexData("extendedtextstrings.txt", "norich-header", FAKE_SID);
            continueBytes = ReadSampleHexData("extendedtextstrings.txt", "norich-continue1", ContinueRecord.sid);
            in1 = TestcaseRecordInputStream.Create(Concat(header, continueBytes));

            strings = new IntMapper<UnicodeString>();
            deserializer = new SSTDeserializer(strings);
            deserializer.ManufactureStrings(1, in1);

            Assert.AreEqual("At a dinner party orAt At At ", strings[0] + "");
        }

        /**
        * Ensure that invalid SST records with an incorrect number of strings specified, does not consume non-continuation records.
        */
        [Test]
        public void TestInvalidSTTRecord()
        {
            byte[] sstRecord = ReadSampleHexData("notenoughstrings.txt", "sst-record", SSTRecord.sid);
            byte[] nonContinuationRecord = ReadSampleHexData("notenoughstrings.txt", "non-continuation-record", ExtSSTRecord.sid);
            RecordInputStream in1 = TestcaseRecordInputStream.Create(Concat(sstRecord, nonContinuationRecord));

            IntMapper<UnicodeString> strings = new IntMapper<UnicodeString>();
            SSTDeserializer deserializer = new SSTDeserializer(strings);

            // The record data in notenoughstrings.txt only contains 1 string, deliberately pass in a larger number.
            deserializer.ManufactureStrings(2, in1);

            Assert.AreEqual("At a dinner party or", strings[0] + "");
            Assert.AreEqual("", strings[1] + "");
        }
    }
}