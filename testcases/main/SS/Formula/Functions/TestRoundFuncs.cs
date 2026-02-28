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

using NUnit.Framework.Constraints;
using NUnit.Framework.Internal;
using System.Numerics;

namespace TestCases.SS.Formula.Functions
{

    using NPOI.SS.Formula.Eval;
    using NUnit.Framework;using NUnit.Framework.Legacy;
    using NPOI.SS.Formula.Functions;
    using System.Numerics;
    using System;

    /**
     * Test cases for ROUND(), ROUNDUP(), ROUNDDOWN()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestRoundFuncs
    {
        // github-43
        // https://github.com/apache/poi/pull/43
        //in java ("ROUNDUP(3987*0.2, 2) currently fails by returning 797.41")
        [Test]
        public void TestRoundUp()
        {
            assertRoundUpEquals(797.40, 3987 * 0.2, 2, 1e-10);
        }
        [Test]
        public void TestRoundDown()
        {
            assertRoundDownEquals(797.40, 3987 * 0.2, 2, 1e-10);
        }
        [Test]
        public void TestRound()
        {
            assertRoundEquals(797.40, 3987 * 0.2, 2, 1e-10);
        }

        //private static NumericFunction F = null;
        [Test]
        public void TestRounddownWithStringArg()
        {

            ValueEval strArg = new StringEval("abc");
            ValueEval[] args = { strArg, new NumberEval(2), };
            ValueEval result = NumericFunction.ROUNDDOWN.Evaluate(args, -1, (short)-1);
            ClassicAssert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }
        [Test]
        public void TestRoundupWithStringArg()
        {

            ValueEval strArg = new StringEval("abc");
            ValueEval[] args = { strArg, new NumberEval(2), };
            ValueEval result = NumericFunction.ROUNDUP.Evaluate(args, -1, (short)-1);
            ClassicAssert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        private static void assertRoundFuncEquals(Function func, double expected, double number, double places, double tolerance)
        {
            ValueEval[] args = { new NumberEval(number), new NumberEval(places), };
            NumberEval result = (NumberEval)func.Evaluate(args, -1, (short)-1);
            ClassicAssert.AreEqual(expected, result.NumberValue, tolerance);
        }

        private static void assertRoundEquals(double expected, double number, double places, double tolerance)
        {
            TestRoundFuncs.assertRoundFuncEquals(NumericFunction.ROUND, expected, number, places, tolerance);
        }

        private static void assertRoundUpEquals(double expected, double number, double places, double tolerance)
        {
            TestRoundFuncs.assertRoundFuncEquals(NumericFunction.ROUNDUP, expected, number, places, tolerance);
        }

        private static void assertRoundDownEquals(double expected, double number, double places, double tolerance)
        {
            TestRoundFuncs.assertRoundFuncEquals(NumericFunction.ROUNDDOWN, expected, number, places, tolerance);
        }
    }
}
