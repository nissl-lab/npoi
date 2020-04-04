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
    using NPOI.SS.Formula;
    using System.IO;
    using System.Collections;

    using NUnit.Framework;
    using NPOI.SS.Formula.PTG;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;

    /**
     * Tests the serialization and deserialization of the FormulaRecord
     * class works correctly.  
     *
     * @author Andrew C. Oliver 
     */
    [TestFixture]
    public class TestFormulaRecord
    {
        [Test]
        public void TestCreateFormulaRecord()
        {
            FormulaRecord record = new FormulaRecord();
            record.Column=((short)0);
            //record.SetRow((short)1);
            record.Row=(1);
            record.XFIndex=((short)4);

            Assert.AreEqual(record.Column, (short)0);
            //Assert.AreEqual(record.Row,(short)1);
            Assert.AreEqual((short)record.Row, (short)1);
            Assert.AreEqual(record.XFIndex, (short)4);
        }

        /**
         * Make sure a NAN value is1 preserved
         * This formula record is1 a representation of =1/0 at row 0, column 0 
         */
        [Test]
        public void TestCheckNanPreserve()
        {
            byte[] formulaByte = {
            0, 0, 0, 0,
            0x0F, 0x00,

			// 8 bytes cached number is a 'special value' in this case
			0x02, // special cached value type 'error'
			0x00,
            FormulaError.DIV0.Code,
            0x00,
            0x00,
            0x00,
            (byte)0xFF,
            (byte)0xFF,

            0x00,
            0x00,
            0x00,
            0x00,

            (byte)0xE0, //18
			(byte)0xFC,
			// Ptgs
			0x07, 0x00, // encoded length
			0x1E, 0x01, 0x00, // IntPtg(1)
			0x1E, 0x00, 0x00, // IntPtg(0)
			0x06, // DividePtg

		};

            FormulaRecord record = new FormulaRecord(TestcaseRecordInputStream.Create(FormulaRecord.sid,  formulaByte));
            Assert.AreEqual(0, record.Row, "Row");
            Assert.AreEqual(0, record.Column, "Column");
            Assert.AreEqual(record.CachedResultType,NPOI.SS.UserModel.CellType.Error);

            byte[] output = record.Serialize();
            Assert.AreEqual(33, output.Length, "Output size"); //includes sid+recordlength

            for (int i = 5; i < 13; i++)
            {
                Assert.AreEqual(formulaByte[i], output[i + 4], "FormulaByte NaN doesn't match");
            }
        }

        /**
         * Tests to see if the shared formula cells properly reSerialize the expPtg
         *
         */
        [Test]
        public void TestExpFormula()
        {
            byte[] formulaByte = new byte[27];

            formulaByte[4] = (byte)0x0F;
            formulaByte[14] = (byte)0x08;
            formulaByte[18] = (byte)0xE0;
            formulaByte[19] = (byte)0xFD;
            formulaByte[20] = (byte)0x05;
            formulaByte[22] = (byte)0x01;
            FormulaRecord record = new FormulaRecord(TestcaseRecordInputStream.Create(FormulaRecord.sid, formulaByte));
            Assert.AreEqual(0, record.Row, "Row");
            Assert.AreEqual(0, record.Column, "Column");
            byte[] output = record.Serialize();
            Assert.AreEqual(31, output.Length, "Output size"); //includes sid+recordlength
            Assert.AreEqual(1, output[26], "OffSet 22");
        }
        [Test]
        public void TestWithConcat()
        {
            // =CHOOSE(2,A2,A3,A4)
            byte[] data = {
				6, 0, 68, 0,
				1, 0, 1, 0, 15, 0, 0, 0, 0, 0, 0, 0, 57,
				64, 0, 0, 12, 0, 12, unchecked((byte)-4), 46, 0, 
				30, 2, 0,	// Int - 2
				25, 4, 3, 0, // Attr
					8, 0, 17, 0, 26, 0, // jumpTable
					35, 0, // chooseOffSet
				36, 1, 0, 0, unchecked((byte)-64), // Ref - A2
				25, 8, 21, 0, // Attr
				36, 2, 0, 0, unchecked((byte)-64), // Ref - A3
				25,	8, 12, 0, // Attr
				36, 3, 0, 0, unchecked((byte)-64), // Ref - A4
				25, 8, 3, 0,  // Attr 
				66, 4, 100, 0 // CHOOSE
		    };
            RecordInputStream inp = new RecordInputStream(new MemoryStream(data));
            inp.NextRecord();

            FormulaRecord fr = new FormulaRecord(inp);

            Ptg[] ptgs = fr.ParsedExpression;
            Assert.AreEqual(9, ptgs.Length);
            Assert.AreEqual(typeof(IntPtg), ptgs[0].GetType());
            Assert.AreEqual(typeof(AttrPtg), ptgs[1].GetType());
            Assert.AreEqual(typeof(RefPtg), ptgs[2].GetType());
            Assert.AreEqual(typeof(AttrPtg), ptgs[3].GetType());
            Assert.AreEqual(typeof(RefPtg), ptgs[4].GetType());
            Assert.AreEqual(typeof(AttrPtg), ptgs[5].GetType());
            Assert.AreEqual(typeof(RefPtg), ptgs[6].GetType());
            Assert.AreEqual(typeof(AttrPtg), ptgs[7].GetType());
            Assert.AreEqual(typeof(FuncVarPtg), ptgs[8].GetType());

            FuncVarPtg choose = (FuncVarPtg)ptgs[8];
            Assert.AreEqual("CHOOSE", choose.Name);
        }
        [Test]
        public void TestReSerialize()
        {
            FormulaRecord formulaRecord = new FormulaRecord();
            formulaRecord.Row = (/*setter*/1);
            formulaRecord.Column = (/*setter*/(short)1);
            formulaRecord.ParsedExpression = (/*setter*/new Ptg[] { new RefPtg("B$5"), });
            formulaRecord.Value = (/*setter*/3.3);
            byte[] ser = formulaRecord.Serialize();
            Assert.AreEqual(31, ser.Length);

            RecordInputStream in1 = TestcaseRecordInputStream.Create(ser);
            FormulaRecord fr2 = new FormulaRecord(in1);
            Assert.AreEqual(3.3, fr2.Value, 0.0);
            Ptg[] ptgs = fr2.ParsedExpression;
            Assert.AreEqual(1, ptgs.Length);
            RefPtg rp = (RefPtg)ptgs[0];
            Assert.AreEqual("B$5", rp.ToFormulaString());
        }

        /**
         * Bug noticed while fixing 46479.  Operands of conditional operator ( ? : ) were swapped
         * inside {@link FormulaRecord}
         */
        [Test]
        public void TestCachedValue_bug46479()
        {
            FormulaRecord fr0 = new FormulaRecord();
            FormulaRecord fr1 = new FormulaRecord();
            // Test some other cached value types 
            fr0.Value = (/*setter*/3.5);
            Assert.AreEqual(3.5, fr0.Value, 0.0);
            fr0.SetCachedResultErrorCode (FormulaError.REF.Code);
            Assert.AreEqual(FormulaError.REF.Code, fr0.CachedErrorValue);

            fr0.SetCachedResultBoolean(false);
            fr1.SetCachedResultBoolean(true);
            if (fr0.CachedBooleanValue == true && fr1.CachedBooleanValue == false)
            {
                throw new AssertionException("Identified bug 46479c");
            }
            Assert.AreEqual(false, fr0.CachedBooleanValue);
            Assert.AreEqual(true, fr1.CachedBooleanValue);
        }
    }
}
