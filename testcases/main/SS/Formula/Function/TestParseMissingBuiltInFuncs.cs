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

namespace TestCases.SS.Formula.Function
{

    using NPOI.HSSF.Model;
    using NPOI.SS.Formula.PTG;
    using NPOI.HSSF.UserModel;
    using System;
    using NUnit.Framework;
    /**
     * Tests parsing of some built-in functions that were not properly
     * registered in POI as of bug #44675, #44733 (March/April 2008).
     * 
     * @author Josh Micich
     */
    [TestFixture]
    public class TestParseMissingBuiltInFuncs
    {

        private static Ptg[] Parse(String formula)
        {
            HSSFWorkbook book = new HSSFWorkbook();
            return HSSFFormulaParser.Parse(formula, book);
        }
        private static void ConfirmFunc(String formula, int expPtgArraySize, bool isVarArgFunc, int funcIx)
        {
            Ptg[] ptgs = Parse(formula);
            Ptg ptgF = ptgs[ptgs.Length - 1];  // func is last RPN token in all these formulas

            // Check critical things in the Ptg array encoding.
            if (!(ptgF is AbstractFunctionPtg))
            {
                throw new Exception("function token missing");
            }
            AbstractFunctionPtg func = (AbstractFunctionPtg)ptgF;
            if (func.FunctionIndex == 255)
            {
                throw new AssertionException("Failed to recognise built-in function in formula '"
                        + formula + "'");
            }
            Assert.AreEqual(expPtgArraySize, ptgs.Length);
            Assert.AreEqual(funcIx, func.FunctionIndex);
            Type expCls = isVarArgFunc ? typeof(FuncVarPtg) : typeof(FuncPtg);
            Assert.AreEqual(expCls, ptgF.GetType());

            // check that Parsed Ptg array Converts back to formula text OK
            HSSFWorkbook book = new HSSFWorkbook();
            String reRenderedFormula = HSSFFormulaParser.ToFormulaString(book, ptgs);
            Assert.AreEqual(formula, reRenderedFormula);
        }
        [Test]
        public void TestDatedif()
        {
            int expSize = 4;   // NB would be 5 if POI Added tAttrVolatile properly
            ConfirmFunc("DATEDIF(NOW(),NOW(),\"d\")", expSize, false, 351);
        }
        [Test]
        public void TestDdb()
        {
            ConfirmFunc("DDB(1,1,1,1,1)", 6, true, 144);
        }
        [Test]
        public void TestAtan()
        {
            ConfirmFunc("ATAN(1)", 2, false, 18);
        }
        [Test]
        public void TestUsdollar()
        {
            ConfirmFunc("USDOLLAR(1)", 2, true, 204);
        }
        [Test]
        public void TestDBCS()
        {
            ConfirmFunc("DBCS(\"abc\")", 2, false, 215);
        }
        [Test]
        public void TestIsnontext()
        {
            ConfirmFunc("ISNONTEXT(\"abc\")", 2, false, 190);
        }
        [Test]
        public void TestDproduct()
        {
            ConfirmFunc("DPRODUCT(C1:E5,\"HarvestYield\",G1:H2)", 4, false, 189);
        }
    }

}