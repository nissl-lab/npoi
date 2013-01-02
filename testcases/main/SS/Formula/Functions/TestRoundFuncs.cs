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

    using NPOI.SS.Formula.Eval;
    using NUnit.Framework;
    using NPOI.SS.Formula.Functions;

    /**
     * Test cases for ROUND(), ROUNDUP(), ROUNDDOWN()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRoundFuncs
    {
        //private static NumericFunction F = null;
        [Test]
        public void TestRounddownWithStringArg()
        {

            ValueEval strArg = new StringEval("abc");
            ValueEval[] args = { strArg, new NumberEval(2), };
            ValueEval result = NumericFunction.ROUNDDOWN.Evaluate(args, -1, (short)-1);
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }
        [Test]
        public void TestRoundupWithStringArg()
        {

            ValueEval strArg = new StringEval("abc");
            ValueEval[] args = { strArg, new NumberEval(2), };
            ValueEval result = NumericFunction.ROUNDUP.Evaluate(args, -1, (short)-1);
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

    }
}
