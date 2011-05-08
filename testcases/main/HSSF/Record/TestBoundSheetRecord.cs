
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
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /**
     * Tests BoundSheetRecord.
     *
     * @see BoundSheetRecord
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestClass]
    public class TestBoundSheetRecord
    {
        public TestBoundSheetRecord()
        {
        }
        [TestMethod]
        public void TestRecordLength()
        {
            BoundSheetRecord record = new BoundSheetRecord("Sheet1");
            Assert.AreEqual(18, record.RecordSize, " 2  +  2  +  4  +   2   +    1     +    1    + len(str)");
        }
        [TestMethod]
        public void TestWideRecordLength()
        {
            BoundSheetRecord record = new BoundSheetRecord("Sheet\u20ac");
            record.Sheetname = ("Sheet\u20ac");

            Assert.AreEqual(24, record.RecordSize, " 2  +  2  +  4  +   2   +    1     +    1    + len(str) * 2");
        }
        [TestMethod]
        public void TestName()
        {
            BoundSheetRecord record = new BoundSheetRecord("1234567890223456789032345678904");

            try
            {
                record.Sheetname = ("s//*s");
                Assert.IsTrue(false, "Should have thrown ArgumentException, but didnt");
            }
            catch (ArgumentException)
            {

            }

        }

    }
}