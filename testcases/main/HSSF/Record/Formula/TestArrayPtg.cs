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

namespace TestCases.HSSF.Record.Formula
{
    using System;
    using NPOI.HSSF.Record.Formula;
    using NPOI.HSSF.Record;

    using TestCases.HSSF;
    using NPOI.HSSF.UserModel;
    using NPOI.Util.IO;

    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /**
     * Tests for <tt>ArrayPtg</tt>
     * 
     * @author Josh Micich
     */
    [TestClass]
    public class TestArrayPtg
    {

        private static byte[] ENCODED_PTG_DATA = {
		0x40, 0x00,
		0x08, 0x00,
		0, 0, 0, 0, 0, 0, 0, 0, 
	};
        private static byte[] ENCODED_CONSTANT_DATA = {
		2,    // 3 columns
		1, 0, // 2 rows
		4, 1, 0, 0, 0, 0, 0, 0, 0, // TRUE
		2, 4, 0, 0, 65, 66, 67, 68, // "ABCD"
		2, 1, 0, 0, 69, // "E"
		1, 0, 0, 0, 0, 0, 0, 0, 0, // 0
		4, 0, 0, 0, 0, 0, 0, 0, 0, // FALSE
		2, 2, 0, 0, 70, 71, // "FG"
	};

        /**
         * Lots of problems with ArrayPtg's encoding of 
         */
        [TestMethod]
        public void TestReadWriteTokenValueBytes()
        {

            ArrayPtg ptg = new ArrayPtg(TestcaseRecordInputStream.CreateWithFakeSid(ENCODED_PTG_DATA));

            ptg.ReadTokenValues(TestcaseRecordInputStream.CreateWithFakeSid(ENCODED_CONSTANT_DATA));
            Assert.AreEqual(3, ptg.ColumnCount);
            Assert.AreEqual(2, ptg.RowCount);
            object[][] values = ptg.GetTokenArrayValues();
            Assert.AreEqual(2, values.Length);


            Assert.AreEqual(true, values[0][0]);
            Assert.AreEqual("ABCD", values[0][1]);
            Assert.AreEqual(0, Convert.ToInt32( values[1][0]));
            Assert.AreEqual(false, values[1][1]);
            Assert.AreEqual("FG", values[1][2]);

            byte[] outBuf = new byte[ENCODED_CONSTANT_DATA.Length];
            ptg.WriteTokenValueBytes(new LittleEndianByteArrayOutputStream(outBuf, 0));

            if (outBuf[0] == 4)
            {
                throw new AssertFailedException("Identified bug 42564b");
            }
            Assert.IsTrue(NPOI.Util.Arrays.Equals(ENCODED_CONSTANT_DATA, outBuf));
        }

        /**
         * Excel stores array elements column by column.  This test makes sure POI does the same.
         */
        [TestMethod]
        public void TestElementOrdering()
        {
            ArrayPtg ptg = new ArrayPtg(TestcaseRecordInputStream.CreateWithFakeSid(ENCODED_PTG_DATA));
            ptg.ReadTokenValues(TestcaseRecordInputStream.CreateWithFakeSid(ENCODED_CONSTANT_DATA));
            Assert.AreEqual(3, ptg.ColumnCount);
            Assert.AreEqual(2, ptg.RowCount);

            Assert.AreEqual(0, ptg.GetValueIndex(0, 0));
            Assert.AreEqual(2, ptg.GetValueIndex(1, 0));
            Assert.AreEqual(4, ptg.GetValueIndex(2, 0));
            Assert.AreEqual(1, ptg.GetValueIndex(0, 1));
            Assert.AreEqual(3, ptg.GetValueIndex(1, 1));
            Assert.AreEqual(5, ptg.GetValueIndex(2, 1));
        }

        /**
         * Test for a bug which was temporarily introduced by the fix for bug 42564.
         * A spReadsheet was added to make the ordering clearer.
         */
        [TestMethod]
        public void TestElementOrderingInSpreadsheet()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex42564-elementOrder.xls");

            // The formula has an array with 3 rows and 5 column 
            String formula = wb.GetSheetAt(0).GetRow(0).GetCell((short)0).CellFormula;
            // TODO - These number literals should not have '.0'. Excel has different number rendering rules

            if (formula.Equals("SUM({1.0,6.0,11.0;2.0,7.0,12.0;3.0,8.0,13.0;4.0,9.0,14.0;5.0,10.0,15.0})"))
            {
                throw new AssertFailedException("Identified bug 42564 b");
            }
            Assert.AreEqual("SUM({1,2,3;4,5,6;7,8,9;10,11,12;13,14,15})", formula);
        }
    }
}