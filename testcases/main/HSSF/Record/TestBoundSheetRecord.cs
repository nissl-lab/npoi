
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
    using System.Collections.Generic;
    using NPOI.Util;

    /**
     * Tests BoundSheetRecord.
     *
     * @see BoundSheetRecord
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    [TestFixture]
    public class TestBoundSheetRecord
    {

        [Test]
        public void TestRecordLength()
        {
            BoundSheetRecord record = new BoundSheetRecord("Sheet1");
            Assert.AreEqual(18, record.RecordSize, " 2  +  2  +  4  +   2   +    1     +    1    + len(str)");
        }
        [Test]
        public void TestWideRecordLength()
        {
            BoundSheetRecord record = new BoundSheetRecord("Sheet\u20ac");
            record.Sheetname = ("Sheet\u20ac");

            Assert.AreEqual(24, record.RecordSize, " 2  +  2  +  4  +   2   +    1     +    1    + len(str) * 2");
        }
        [Test]
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
        [Test]
        public void TestDeSerializeUnicode() {

		byte[] data = HexRead.ReadFromString(""
			+ "85 00 1A 00" // sid, length
			+ "3C 09 00 00" // bof
			+ "00 00"// flags
			+ "09 01" // str-len. unicode flag
			// string data
			+ "21 04 42 04 40 04"
			+ "30 04 3D 04 38 04"
			+ "47 04 3A 04 30 04"
		);

		RecordInputStream in1 = TestcaseRecordInputStream.Create(data);
		BoundSheetRecord bsr = new BoundSheetRecord(in1);
		// sheet name is unicode Russian for 'minor page'
		Assert.AreEqual("\u0421\u0442\u0440\u0430\u043D\u0438\u0447\u043A\u0430", bsr.Sheetname);

		byte[] data2 = bsr.Serialize();
		Assert.IsTrue(Arrays.Equals(data, data2));
	}
        [Test]
        public void TestOrdering()
        {
            BoundSheetRecord bs1 = new BoundSheetRecord("SheetB");
            BoundSheetRecord bs2 = new BoundSheetRecord("SheetC");
            BoundSheetRecord bs3 = new BoundSheetRecord("SheetA");
            bs1.PositionOfBof = (/*setter*/11);
            bs2.PositionOfBof = (/*setter*/33);
            bs3.PositionOfBof = (/*setter*/22);

            List<BoundSheetRecord> l = new List<BoundSheetRecord>();
            l.Add(bs1);
            l.Add(bs2);
            l.Add(bs3);

            BoundSheetRecord[] r = BoundSheetRecord.OrderByBofPosition(l);
            Assert.AreEqual(3, r.Length);
            Assert.AreEqual(bs1, r[0]);
            Assert.AreEqual(bs3, r[1]);
            Assert.AreEqual(bs2, r[2]);
        }
        [Test]
        public void TestValidNames()
        {
            ConfirmValid("Sheet1", true);
            ConfirmValid("O'Brien's sales", true);
            ConfirmValid(" data # ", true);
            ConfirmValid("data $1.00", true);

            ConfirmValid("data?", false);
            ConfirmValid("abc/def", false);
            ConfirmValid("data[0]", false);
            ConfirmValid("data*", false);
            ConfirmValid("abc\\def", false);
            ConfirmValid("'data", false);
            ConfirmValid("data'", false);
        }

        private static void ConfirmValid(String sheetName, bool expectedResult)
        {

            try
            {
                new BoundSheetRecord(sheetName);
                if (!expectedResult)
                {
                    throw new AssertionException("Expected sheet name '" + sheetName + "' to be invalid");
                }
            }
            catch (ArgumentException)
            {
                if (expectedResult)
                {
                    throw new AssertionException("Expected sheet name '" + sheetName + "' to be valid");
                }
            }
        }

    }
}