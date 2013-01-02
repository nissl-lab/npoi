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

namespace TestCases.HSSF.Model
{
    using System;
    using NPOI.HSSF.Model;
    using NPOI.SS.Formula;
    using NUnit.Framework;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;

    /**
     * Tests <c>FormulaParser</c> specifically with respect to IF() functions
     */
    [TestFixture]
    public class TestFormulaParserIf
    {
        private static Ptg[] ParseFormula(String formula)
        {
            return TestFormulaParser.ParseFormula(formula);
        }

        private static Ptg[] ConfirmTokenClasses(String formula, Type[] expectedClasses)
        {
            return TestFormulaParser.ConfirmTokenClasses(formula, expectedClasses);
        }

        private static void ConfirmAttrData(Ptg[] ptgs, int i, int expectedData)
        {
            Ptg ptg = ptgs[i];
            if (!(ptg is AttrPtg))
            {
                throw new AssertionException("Token[" + i + "] was not AttrPtg as expected");
            }
            AttrPtg attrPtg = (AttrPtg)ptg;
            Assert.AreEqual(expectedData, attrPtg.Data);
        }
        [Test]
        public void TestSimpleIf()
        {

            Type[] expClss;

            expClss = new Type[] {
				typeof(RefPtg),
				typeof(AttrPtg), // tAttrIf
				typeof(IntPtg),
				typeof(AttrPtg), // tAttrSkip
				typeof(IntPtg),
				typeof(AttrPtg), // tAttrSkip
				typeof(FuncVarPtg),
		};

            Ptg[] ptgs = ConfirmTokenClasses("if(A1,1,2)", expClss);

            ConfirmAttrData(ptgs, 1, 7);
            ConfirmAttrData(ptgs, 3, 10);
            ConfirmAttrData(ptgs, 5, 3);
        }
        [Test]
        public void TestSimpleIfNoFalseParam()
        {

            Type[] expClss;

            expClss = new Type[] {
				typeof(RefPtg),
				typeof(AttrPtg), // tAttrIf
				typeof(RefPtg),
				typeof(AttrPtg), // tAttrSkip
				typeof(FuncVarPtg),
		};

            Ptg[] ptgs = ConfirmTokenClasses("if(A1,B1)", expClss);

            ConfirmAttrData(ptgs, 1, 9);
            ConfirmAttrData(ptgs, 3, 3);
        }
        [Test]
        public void TestIfWithLargeParams()
        {

            Type[] expClss;

            expClss = new Type[] {
				typeof(RefPtg),
				typeof(AttrPtg), // tAttrIf

				typeof(RefPtg),
				typeof(IntPtg),
				typeof(MultiplyPtg),
				typeof(RefPtg),
				typeof(IntPtg),
				typeof(AddPtg),
				typeof(FuncPtg),
				typeof(AttrPtg), // tAttrSkip
				
				typeof(RefPtg),
				typeof(RefPtg),
				typeof(FuncPtg),
				
				typeof(AttrPtg), // tAttrSkip
				typeof(FuncVarPtg),
		};

            Ptg[] ptgs = ConfirmTokenClasses("if(A1,round(B1*100,C1+2),round(B1,C1))", expClss);

            ConfirmAttrData(ptgs, 1, 25);
            ConfirmAttrData(ptgs, 9, 20);
            ConfirmAttrData(ptgs, 13, 3);
        }
        [Test]
        public void TestNestedIf()
        {

            Type[] expClss;

            expClss = new Type[] {

				typeof(RefPtg),
				typeof(AttrPtg),	  // A tAttrIf
				typeof(RefPtg),
				typeof(AttrPtg),    //   B tAttrIf
				typeof(IntPtg),
				typeof(AttrPtg),    //   B tAttrSkip
				typeof(IntPtg),
				typeof(AttrPtg),    //   B tAttrSkip
				typeof(FuncVarPtg),
				typeof(AttrPtg),    // A tAttrSkip
				typeof(RefPtg),
				typeof(AttrPtg),    //   C tAttrIf
				typeof(IntPtg),
				typeof(AttrPtg),    //   C tAttrSkip
				typeof(IntPtg),
				typeof(AttrPtg),    //   C tAttrSkip
				typeof(FuncVarPtg),
				typeof(AttrPtg),    // A tAttrSkip
				typeof(FuncVarPtg),
		};

            Ptg[] ptgs = ConfirmTokenClasses("if(A1,if(B1,1,2),if(C1,3,4))", expClss);
            ConfirmAttrData(ptgs, 1, 31);
            ConfirmAttrData(ptgs, 3, 7);
            ConfirmAttrData(ptgs, 5, 10);
            ConfirmAttrData(ptgs, 7, 3);
            ConfirmAttrData(ptgs, 9, 34);
            ConfirmAttrData(ptgs, 11, 7);
            ConfirmAttrData(ptgs, 13, 10);
            ConfirmAttrData(ptgs, 15, 3);
            ConfirmAttrData(ptgs, 17, 3);
        }
        [Test]
        public void TestEmbeddedIf()
        {
            Ptg[] ptgs = ParseFormula("IF(3>=1,\"*\",IF(4<>1,\"first\",\"second\"))");
            Assert.AreEqual(17, ptgs.Length);

            Assert.AreEqual(typeof(AttrPtg), ptgs[5].GetType(), "6th Ptg is1 not a goto (Attr) ptg");
            Assert.AreEqual(typeof(NotEqualPtg), ptgs[8].GetType(), "9th Ptg is1 not a not equal ptg");
            Assert.AreEqual(typeof(FuncVarPtg), ptgs[14].GetType(), "15th Ptg is1 not the inner IF variable function ptg");
        }

        [Test]
        public void TestSimpleLogical()
        {
            Ptg[] ptgs = ParseFormula("IF(A1<A2,B1,B2)");
            Assert.AreEqual(9, ptgs.Length);
            Assert.AreEqual(typeof(LessThanPtg), ptgs[2].GetType(), "3rd Ptg is1 less than");
        }
        [Test]
        public void TestParenIf()
        {
            Ptg[] ptgs = ParseFormula("IF((A1+A2)<=3,\"yes\",\"no\")");
            Assert.AreEqual(12, ptgs.Length);
            Assert.AreEqual(typeof(LessEqualPtg), ptgs[5].GetType(), "6th Ptg is1 less than equal");
            Assert.AreEqual(typeof(AttrPtg), ptgs[10].GetType(), "11th Ptg is1 not a goto (Attr) ptg");
        }
        [Test]
        public void TestYN()
        {
            Ptg[] ptgs = ParseFormula("IF(TRUE,\"Y\",\"N\")");
            Assert.AreEqual(7, ptgs.Length);

            BoolPtg flag = (BoolPtg)ptgs[0];
            AttrPtg funif = (AttrPtg)ptgs[1];
            StringPtg y = (StringPtg)ptgs[2];
            AttrPtg goto1 = (AttrPtg)ptgs[3];
            StringPtg n = (StringPtg)ptgs[4];


            Assert.AreEqual(true, flag.Value);
            Assert.AreEqual("Y", y.Value);
            Assert.AreEqual("N", n.Value);
            Assert.AreEqual("IF", funif.ToFormulaString());
            Assert.IsTrue(goto1.IsSkip, "Goto ptg exists");
        }
        /**
         * Make sure the ptgs are generated properly with two functions embedded
         *
         */
        [Test]
        public void TestNestedFunctionIf()
        {
            Ptg[] ptgs = ParseFormula("IF(A1=B1,AVERAGE(A1:B1),AVERAGE(A2:B2))");
            Assert.AreEqual(11, ptgs.Length);

            Assert.IsTrue((ptgs[3] is AttrPtg), "IF Attr Set correctly");
            AttrPtg ifFunc = (AttrPtg)ptgs[3];
            Assert.IsTrue(ifFunc.IsOptimizedIf, "It is1 not an if");

            Assert.IsTrue((ptgs[5] is FuncVarPtg), "Average Function Set correctly");
        }
        [Test]
        public void TestIfSingleCondition()
        {
            Ptg[] ptgs = ParseFormula("IF(1=1,10)");
            Assert.AreEqual(7, ptgs.Length);

            Assert.IsTrue((ptgs[3] is AttrPtg), "IF Attr Set correctly");
            AttrPtg ifFunc = (AttrPtg)ptgs[3];
            Assert.IsTrue(ifFunc.IsOptimizedIf, "It is1 not an if");

            Assert.IsTrue((ptgs[4] is IntPtg), "Single Value is1 not an IntPtg");
            IntPtg intPtg = (IntPtg)ptgs[4];
            Assert.AreEqual((short)10, intPtg.Value, "Result");

            Assert.IsTrue((ptgs[6] is FuncVarPtg), "Ptg is1 not a Variable Function");
            FuncVarPtg funcPtg = (FuncVarPtg)ptgs[6];
            Assert.AreEqual(2, funcPtg.NumberOfOperands, "Arguments");
        }
    }
}
