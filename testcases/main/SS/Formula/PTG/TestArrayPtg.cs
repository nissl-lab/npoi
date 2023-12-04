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

namespace TestCases.SS.Formula.PTG
{
    using System;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;
    using NPOI.Util;

    using TestCases.HSSF;
    using TestCases.HSSF.Record;
    /**
     * Tests for <c>ArrayPtg</c>
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestArrayPtg
    {

        private static byte[] ENCODED_PTG_DATA = {
		0x40,
		0, 0, 0, 0, 0, 0, 0,
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

        private static ArrayPtg Create(byte[] initialData, byte[] constantData)
        {
            ArrayPtg.Initial ptgInit = new ArrayPtg.Initial(TestcaseRecordInputStream.CreateLittleEndian(initialData));
            return ptgInit.FinishReading(TestcaseRecordInputStream.CreateLittleEndian(constantData));
        }

        /**
         * Lots of problems with ArrayPtg's decoding and encoding of the element value data
         */
        [Test]
        [Ignore("TODO FIX CI TESTS")]
        public void TestReadWriteTokenValueBytes()
        {
            ArrayPtg ptg = Create(ENCODED_PTG_DATA, ENCODED_CONSTANT_DATA);
            Assert.AreEqual(3, ptg.ColumnCount);
            Assert.AreEqual(2, ptg.RowCount);
            Object[,] values = ptg.GetTokenArrayValues();
            Assert.AreEqual(2, values.Length);


            Assert.AreEqual(true, values[0,0]);
            Assert.AreEqual("ABCD", values[0,1]);
            Assert.AreEqual(0d, values[1,0]);
            Assert.AreEqual(false, values[1,1]);
            Assert.AreEqual("FG", values[1,2]);

            byte[] outBuf = new byte[ENCODED_CONSTANT_DATA.Length];
            ptg.WriteTokenValueBytes(new LittleEndianByteArrayOutputStream(outBuf, 0));

            if (outBuf[0] == 4)
            {
                throw new AssertionException("Identified bug 42564b");
            }
            Assert.IsTrue(Arrays.Equals(ENCODED_CONSTANT_DATA, outBuf));
        }


        /**
         * Excel stores array elements column by column.  This Test Makes sure POI does the same.
         */
        [Test]
        public void TestElementOrdering()
        {
            ArrayPtg ptg = Create(ENCODED_PTG_DATA, ENCODED_CONSTANT_DATA);
            Assert.AreEqual(3, ptg.ColumnCount);
            Assert.AreEqual(2, ptg.RowCount);

            Assert.AreEqual(0, ptg.GetValueIndex(0, 0));
            Assert.AreEqual(1, ptg.GetValueIndex(1, 0));
            Assert.AreEqual(2, ptg.GetValueIndex(2, 0));
            Assert.AreEqual(3, ptg.GetValueIndex(0, 1));
            Assert.AreEqual(4, ptg.GetValueIndex(1, 1));
            Assert.AreEqual(5, ptg.GetValueIndex(2, 1));
        }

        /**
         * Test for a bug which was temporarily introduced by the fix for bug 42564.
         * A spreadsheet was Added to make the ordering Clearer.
         */
        [Test]
        public void TestElementOrderingInSpreadsheet()
        {
            HSSFWorkbook wb = HSSFTestDataSamples.OpenSampleWorkbook("ex42564-elementOrder.xls");

            // The formula has an array with 3 rows and 5 columns
            String formula = wb.GetSheetAt(0).GetRow(0).GetCell(0).CellFormula;

            if (formula.Equals("SUM({1,6,11;2,7,12;3,8,13;4,9,14;5,10,15})"))
            {
                throw new AssertionException("Identified bug 42564 b");
            }
            Assert.AreEqual("SUM({1,2,3,4,5;6,7,8,9,10;11,12,13,14,15})", formula);
        }
        [Test]
        public void TestToFormulaString()
        {
            ArrayPtg ptg = Create(ENCODED_PTG_DATA, ENCODED_CONSTANT_DATA);
            String actualFormula;
            try
            {
                actualFormula = ptg.ToFormulaString();
            }
            catch (ArgumentException e)
            {
                if (e.Message.Equals("Unexpected constant class (java.lang.Boolean)"))
                {
                    throw new AssertionException("Identified bug 45380");
                }
                throw e;
            }
            Assert.AreEqual("{TRUE,\"ABCD\",\"E\";0,FALSE,\"FG\"}", actualFormula);
        }

        /**
         * worth Checking since AttrPtg.sid=0x20 and Ptg.CLASS_* = (0x00, 0x20, and 0x40)
         */
        [Test]
        public void TestOperandClassDecoding()
        {
            ConfirmOperandClassDecoding(Ptg.CLASS_REF);
            ConfirmOperandClassDecoding(Ptg.CLASS_VALUE);
            ConfirmOperandClassDecoding(Ptg.CLASS_ARRAY);
        }

        private static void ConfirmOperandClassDecoding(byte operandClass)
        {
            byte[] fullData = concat(ENCODED_PTG_DATA, ENCODED_CONSTANT_DATA);

            // Force encoded operand class for tArray
            fullData[0] = (byte)(ArrayPtg.sid + operandClass);

            ILittleEndianInput in1 = TestcaseRecordInputStream.CreateLittleEndian(fullData);

            Ptg[] ptgs = Ptg.ReadTokens(ENCODED_PTG_DATA.Length, in1);
            Assert.AreEqual(1, ptgs.Length);
            ArrayPtg aPtg = (ArrayPtg)ptgs[0];
            Assert.AreEqual(operandClass, aPtg.PtgClass);
        }

        private static byte[] concat(byte[] a, byte[] b)
        {
            byte[] result = new byte[a.Length + b.Length];
            Array.Copy(a, 0, result, 0, a.Length);
            Array.Copy(b, 0, result, a.Length, b.Length);
            return result;
        }
    }

}