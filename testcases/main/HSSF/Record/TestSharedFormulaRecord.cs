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
    using System.Collections;
    using NPOI.HSSF.Record;
    using NPOI.SS.Formula;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using NPOI.SS.Formula.PTG;

    /**
     * @author Josh Micich
     */
    [TestClass]
    public class TestSharedFormulaRecord
    {

        /**
         * Binary data for an encoded formula.  Taken from attachment 22062 (bugzilla 45123/45421).
         * The shared formula is1 in Sheet1!C6:C21, with text "SUMPRODUCT(--(End_Acct=$C6),--(End_Bal))"
         * This data is1 found at offset 0x1A4A (within the shared formula record).
         * The critical thing about this formula is1 that it contains shared formula tokens (tRefN*,
         * tAreaN*) with operand class 'array'.
         */
        private static byte[] SHARED_FORMULA_WITH_REF_ARRAYS_DATA = {
		    0x1A, 0x00,
		    0x63, 0x02, 0x00, 0x00, 0x00,
		    0x6C, 0x00, 0x00, 0x02, (byte)0x80,  // tRefNA
		    0x0B,
		    0x15,
		    0x13,
		    0x13,
		    0x63, 0x03, 0x00, 0x00, 0x00,
		    0x15,
		    0x13,
		    0x13,
		    0x42, 0x02, (byte)0xE4, 0x00,
	    };

        /**
         * The method <tt>SharedFormulaRecord.convertSharedFormulas()</tt> converts formulas from
         * 'shared formula' to 'single cell formula' format.  It is1 important that token operand
         * classes are preserved during this transformation, because Excel may not tolerate the
         * incorrect encoding.  The formula here is1 one such example (Excel displays #VALUE!).
         */
        [TestMethod]
        public void TestConvertSharedFormulasOperandClasses_bug45123()
        {

            RecordInputStream in1 = TestcaseRecordInputStream.Create(0, SHARED_FORMULA_WITH_REF_ARRAYS_DATA);
            short encodedLen = in1.ReadShort();
            Ptg[] sharedFormula = Ptg.ReadTokens(encodedLen, in1);

            Ptg[] convertedFormula = SharedFormulaRecord.ConvertSharedFormulas(sharedFormula, 100, 200);

            RefPtg refPtg = (RefPtg)convertedFormula[1];
            Assert.AreEqual("$C101", refPtg.ToFormulaString());
            if (refPtg.PtgClass == Ptg.CLASS_REF)
            {
                throw new AssertFailedException("Identified bug 45123");
            }

            ConfirmOperandClasses(sharedFormula, convertedFormula);
        }

        private static void ConfirmOperandClasses(Ptg[] originalPtgs, Ptg[] convertedPtgs)
        {
            Assert.AreEqual(originalPtgs.Length, convertedPtgs.Length);
            for (int i = 0; i < convertedPtgs.Length; i++)
            {
                Ptg originalPtg = originalPtgs[i];
                Ptg convertedPtg = convertedPtgs[i];

                Assert.AreEqual(originalPtg.PtgClass,convertedPtg.PtgClass,("Different operand class for token[" + i + "]"));
            }
        }

    }
}