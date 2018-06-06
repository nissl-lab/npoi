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

namespace TestCases.HSSF.Model
{
    using System;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;
    using NPOI.HSSF.Model;

    /**
     * Tests specific formula examples in <tt>OperandClassTransformer</tt>.
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestOperandClassTransformer
    {

        private static Ptg[] ParseFormula(String formula)
        {
            Ptg[] result = HSSFFormulaParser.Parse(formula, (HSSFWorkbook)null);
            Assert.IsNotNull(result, "Ptg array should not be null");
            return result;
        }

        [Test]
        public void TestMdeterm()
        {
            String formula = "MDETERM(ABS(A1))";
            Ptg[] ptgs = ParseFormula(formula);

            ConfirmTokenClass(ptgs, 0, Ptg.CLASS_ARRAY);
            ConfirmFuncClass(ptgs, 1, "ABS", Ptg.CLASS_ARRAY);
            ConfirmFuncClass(ptgs, 2, "MDETERM", Ptg.CLASS_VALUE);
        }

        /**
         * In the example: <code>INDEX(PI(),1)</code>, Excel encodes PI() as 'array'.  It is not clear
         * what rule justifies this. POI currently encodes it as 'value' which Excel(2007) seems to 
         * tolerate. Changing the metadata for INDEX to have first parameter as 'array' class breaks 
         * other formulas involving INDEX.  It seems like a special case needs to be made.  Perhaps an 
         * important observation is that INDEX is one of very few functions that returns 'reference' type.
         * 
         * This test has been Added but disabled in order to document this issue.
         */
        [Ignore("this test is disabled in poi.")]//
        public void DISABLED_TestIndexPi1()
        {
            String formula = "INDEX(PI(),1)";
            Ptg[] ptgs = ParseFormula(formula);

            ConfirmFuncClass(ptgs, 1, "PI", Ptg.CLASS_ARRAY); // fails as of POI 3.1
            ConfirmFuncClass(ptgs, 2, "INDEX", Ptg.CLASS_VALUE);
        }

        /**
         * Even though count expects args of type R, because A1 is a direct operand of a
         * value operator it must Get type V
         */
        [Test]
        public void TestDirectOperandOfValueOperator()
        {
            String formula = "COUNT(A1*1)";
            Ptg[] ptgs = ParseFormula(formula);
            if (ptgs[0].PtgClass == Ptg.CLASS_REF)
            {
                throw new AssertionException("Identified bug 45348");
            }

            ConfirmTokenClass(ptgs, 0, Ptg.CLASS_VALUE);
            ConfirmTokenClass(ptgs, 3, Ptg.CLASS_VALUE);
        }

        /**
         * A cell ref passed to a function expecting type V should be Converted to type V
         */
        [Test]
        public void TestRtoV()
        {

            String formula = "Lookup(A1, A3:A52, B3:B52)";
            Ptg[] ptgs = ParseFormula(formula);
            ConfirmTokenClass(ptgs, 0, Ptg.CLASS_VALUE);
        }

        [Test]
        public void TestComplexIRR_bug45041()
        {
            String formula = "(1+IRR(SUMIF(A:A,ROW(INDIRECT(MIN(A:A)&\":\"&MAX(A:A))),B:B),0))^365-1";
            Ptg[] ptgs = ParseFormula(formula);

            FuncVarPtg rowFunc = (FuncVarPtg)ptgs[10];
            FuncVarPtg sumifFunc = (FuncVarPtg)ptgs[12];
            Assert.AreEqual("ROW", rowFunc.Name);
            Assert.AreEqual("SUMIF", sumifFunc.Name);

            if (rowFunc.PtgClass == Ptg.CLASS_VALUE || sumifFunc.PtgClass == Ptg.CLASS_VALUE)
            {
                throw new AssertionException("Identified bug 45041");
            }
            ConfirmTokenClass(ptgs, 1, Ptg.CLASS_REF);
            ConfirmTokenClass(ptgs, 2, Ptg.CLASS_REF);
            ConfirmFuncClass(ptgs, 3, "MIN", Ptg.CLASS_VALUE);
            ConfirmTokenClass(ptgs, 6, Ptg.CLASS_REF);
            ConfirmFuncClass(ptgs, 7, "MAX", Ptg.CLASS_VALUE);
            ConfirmFuncClass(ptgs, 9, "INDIRECT", Ptg.CLASS_REF);
            ConfirmFuncClass(ptgs, 10, "ROW", Ptg.CLASS_ARRAY);
            ConfirmTokenClass(ptgs, 11, Ptg.CLASS_REF);
            ConfirmFuncClass(ptgs, 12, "SUMIF", Ptg.CLASS_ARRAY);
            ConfirmFuncClass(ptgs, 14, "IRR", Ptg.CLASS_VALUE);
        }

        private void ConfirmFuncClass(Ptg[] ptgs, int i, String expectedFunctionName, byte operandClass)
        {
            ConfirmTokenClass(ptgs, i, operandClass);
            AbstractFunctionPtg afp = (AbstractFunctionPtg)ptgs[i];
            Assert.AreEqual(expectedFunctionName, afp.Name);
        }

        private void ConfirmTokenClass(Ptg[] ptgs, int i, byte operandClass)
        {
            Ptg ptg = ptgs[i];
            if (ptg.IsBaseToken)
            {
                throw new AssertionException("ptg[" + i + "] is a base token");
            }
            if (operandClass != ptg.PtgClass)
            {
                throw new AssertionException("Wrong operand class for ptg ("
                        + ptg.ToString() + "). Expected " + GetOperandClassName(operandClass)
                        + " but got " + GetOperandClassName(ptg.PtgClass));
            }
        }

        private static String GetOperandClassName(byte ptgClass)
        {
            switch (ptgClass)
            {
                case Ptg.CLASS_REF:
                    return "R";
                case Ptg.CLASS_VALUE:
                    return "V";
                case Ptg.CLASS_ARRAY:
                    return "A";
            }
            throw new Exception("Unknown operand class (" + ptgClass + ")");
        }
    }

}