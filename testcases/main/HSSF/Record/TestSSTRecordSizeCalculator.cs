
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
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.Util.Collections;
    using NPOI.HSSF.Record;
    using NPOI.HSSF.Record.Cont;


    /**
     * Tests that records size calculates correctly.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestClass]
    public class TestSSTRecordSizeCalculator
    {
        private static String SMALL_STRING = "Small string";
        private static int COMPRESSED_PLAIN_STRING_OVERHEAD = 3;
        //    private List recordLengths;
        private IntMapper strings;
        private static int OPTION_FIELD_SIZE = 1;

        public TestSSTRecordSizeCalculator()
        {

        }
        [TestMethod]
        public void TestBasic()
        {
            strings.Add(MakeUnicodeString(SMALL_STRING));
            ConfirmSize(SSTRecord.SST_RECORD_OVERHEAD
                    + COMPRESSED_PLAIN_STRING_OVERHEAD
                    + SMALL_STRING.Length);
        }
        [TestMethod]
        public void TestBigStringAcrossUnicode()
        {
            String bigString = new String(new char[SSTRecord.MAX_DATA_SPACE + 100]);
            strings.Add(MakeUnicodeString(bigString));
            ConfirmSize(SSTRecord.SST_RECORD_OVERHEAD
                    + COMPRESSED_PLAIN_STRING_OVERHEAD
                    + SSTRecord.MAX_DATA_SPACE
                    + SSTRecord.STD_RECORD_OVERHEAD
                    + OPTION_FIELD_SIZE
                    + 100);
        }
        [TestMethod]
        public void TestPerfectFit()
        {
            int perfectFit = SSTRecord.MAX_DATA_SPACE - COMPRESSED_PLAIN_STRING_OVERHEAD;
            strings.Add(MakeUnicodeString(perfectFit));
            ConfirmSize(SSTRecord.SST_RECORD_OVERHEAD
                    + COMPRESSED_PLAIN_STRING_OVERHEAD
                    + perfectFit);
        }
        [TestMethod]
        public void TestJustOversized()
        {
            int tooBig = SSTRecord.MAX_DATA_SPACE - COMPRESSED_PLAIN_STRING_OVERHEAD + 1;
            strings.Add(MakeUnicodeString(tooBig));
            ConfirmSize(SSTRecord.SST_RECORD_OVERHEAD
                    + COMPRESSED_PLAIN_STRING_OVERHEAD
                    + tooBig - 1
                // continue record
                    + SSTRecord.STD_RECORD_OVERHEAD
                    + OPTION_FIELD_SIZE + 1);
        }
        [TestMethod]
        public void TestSecondStringStartsOnNewContinuation()
        {
            int perfectFit = SSTRecord.MAX_DATA_SPACE - COMPRESSED_PLAIN_STRING_OVERHEAD;
            strings.Add(MakeUnicodeString(perfectFit));
            strings.Add(MakeUnicodeString(SMALL_STRING));
            ConfirmSize(SSTRecord.SST_RECORD_OVERHEAD
                    + SSTRecord.MAX_DATA_SPACE
                // second string
                    + SSTRecord.STD_RECORD_OVERHEAD
                    + COMPRESSED_PLAIN_STRING_OVERHEAD
                    + SMALL_STRING.Length);
        }
        [TestMethod]
        public void TestHeaderCrossesNormalContinuePoint()
        {
            int almostPerfectFit = SSTRecord.MAX_DATA_SPACE - COMPRESSED_PLAIN_STRING_OVERHEAD - 2;
            strings.Add(MakeUnicodeString(almostPerfectFit));
            String oneCharString = new String(new char[1]);
            strings.Add(MakeUnicodeString(oneCharString));
            ConfirmSize(SSTRecord.SST_RECORD_OVERHEAD
                    + COMPRESSED_PLAIN_STRING_OVERHEAD
                    + almostPerfectFit
                // second string
                    + SSTRecord.STD_RECORD_OVERHEAD
                    + COMPRESSED_PLAIN_STRING_OVERHEAD
                    + oneCharString.Length);

        }

        [TestInitialize]
        public void SetUp()
        {
            strings = new IntMapper();
        }

        private static UnicodeString MakeUnicodeString(int size)
        {
            String s = new String(new char[size]);
            return MakeUnicodeString(s);
        }
        private static UnicodeString MakeUnicodeString(String s)
        {
            UnicodeString st = new UnicodeString(s);
            st.OptionFlags = ((byte)0);
            return st;
        }
        private void ConfirmSize(int expectedSize)
        {
            ContinuableRecordOutput cro = ContinuableRecordOutput.CreateForCountingOnly();
            SSTSerializer ss = new SSTSerializer(strings, 0, 0);
            ss.Serialize(cro);
            Assert.AreEqual(expectedSize, cro.TotalSize);
        }

    }
}