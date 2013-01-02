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
     * Tests for {@link StyleRecord}
     */
    [TestFixture]
    public class TestStyleRecord
    {
        [Test]
        public void TestUnicodeReadName()
        {
            byte[] data = HexRead.ReadFromString(
                    "11 00 09 00 01 38 5E C4 89 5F 00 53 00 68 00 65 00 65 00 74 00 31 00");
            RecordInputStream in1 = TestcaseRecordInputStream.Create(StyleRecord.sid, data);
            StyleRecord sr = new StyleRecord(in1);
            Assert.AreEqual("\u5E38\u89C4_Sheet1", sr.Name); // "<Conventional>_Sheet1"
            byte[] ser;
            try
            {
                ser = sr.Serialize();
            }
            catch (InvalidOperationException e)
            {
                if (e.Message.Equals("Incorrect number of bytes written - expected 27 but got 18"))
                {
                    throw new AssertionException("Identified bug 46385");
                }
                throw e;
            }
            TestcaseRecordInputStream.ConfirmRecordEncoding(StyleRecord.sid, data, ser);
        }
    }

}