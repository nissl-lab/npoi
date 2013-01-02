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

namespace TestCases.SS.Formula.Functions
{
    using System;
    using NPOI.SS.Formula.Eval;
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;

    /**
     * 
     * Test cases for {@link TextFunction#LEFT} and {@link TextFunction#RIGHT}
     * 
     * @author Brendan Nolan
     *
     */
    [TestFixture]
    public class TestLeftRight
    {

        private static NumberEval NEGATIVE_OPERAND = new NumberEval(-1.0);
        private static StringEval ANY_STRING_VALUE = new StringEval("ANYSTRINGVALUE");


        private static ValueEval invokeLeft(ValueEval text, ValueEval operand)
        {
            ValueEval[] args = new ValueEval[] { text, operand };
            return TextFunction.LEFT.Evaluate(args, -1, (short)-1);
        }

        private static ValueEval invokeRight(ValueEval text, ValueEval operand)
        {
            ValueEval[] args = new ValueEval[] { text, operand };
            return TextFunction.RIGHT.Evaluate(args, -1, (short)-1);
        }
        [Test]
        public void TestLeftRight_bug49841()
        {

            try
            {
                invokeLeft(ANY_STRING_VALUE, NEGATIVE_OPERAND);
                invokeRight(ANY_STRING_VALUE, NEGATIVE_OPERAND);
            }
            catch (IndexOutOfRangeException)
            {
                Assert.Fail("Identified bug 49841");
            }

        }
        [Test]
        public void TestLeftRightNegativeOperand()
        {

            Assert.AreEqual(ErrorEval.VALUE_INVALID, invokeRight(ANY_STRING_VALUE, NEGATIVE_OPERAND));
            Assert.AreEqual(ErrorEval.VALUE_INVALID, invokeLeft(ANY_STRING_VALUE, NEGATIVE_OPERAND));

        }

    }

}