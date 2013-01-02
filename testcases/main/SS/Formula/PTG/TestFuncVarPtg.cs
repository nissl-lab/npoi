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
    using NPOI.HSSF.Model;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.Formula.PTG;

    /**
     * @author Josh Micich
     */
    [TestFixture]
    public class TestFuncVarPtg
    {

        /**
         * The first fix for bugzilla 44675 broke the encoding of SUM formulas (and probably others).
         * The operand classes of the parameters to SUM() should be Coerced to 'reference' not 'value'.
         * In the case of SUM, Excel Evaluates the formula to '#VALUE!' if a parameter operand class is
         * wrong.  In other cases Excel seems to tolerate bad operand classes.</p>
         * This functionality is related to the SetParameterRVA() methods of <c>FormulaParser</c>
         */
        [Test]
        public void TestOperandClass()
        {
            HSSFWorkbook book = new HSSFWorkbook();
            Ptg[] ptgs = HSSFFormulaParser.Parse("sum(A1:A2)", book);
            Assert.AreEqual(2, ptgs.Length);
            Assert.AreEqual(typeof(AreaPtg), ptgs[0].GetType());

            switch (ptgs[0].PtgClass)
            {
                case Ptg.CLASS_REF:
                    // correct behaviour
                    break;
                case Ptg.CLASS_VALUE:
                    throw new AssertionException("Identified bug 44675b");
                default:
                    throw new Exception("Unexpected operand class");
            }
        }
    }

}