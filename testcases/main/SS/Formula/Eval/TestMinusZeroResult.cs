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

namespace TestCases.SS.Formula.Eval
{
    using System;
    using NUnit.Framework;
    using NPOI.SS.Formula.Eval;
    using NPOI.SS.Formula.Functions;
    using NPOI.Util;
    using TestCases.Exceptions;

    /**
     * IEEE 754 defines a quantity '-0.0' which is distinct from '0.0'.
     * Negative zero is not easy to observe in Excel, since it is usually Converted to 0.0.
     * (Note - the results of XLL Add-in functions don't seem to be Converted, so they are one
     * reliable avenue to observe Excel's treatment of '-0.0' as an operand.)
     * <p/>
     * POI attempts to emulate Excel faithfully, so this class Tests
     * two aspects of '-0.0' in formula Evaluation:
     * <ol>
     * <li>For most operation results '-0.0' is Converted to '0.0'.</li>
     * <li>Comparison operators have slightly different rules regarding '-0.0'.</li>
     * </ol>
     * @author Josh Micich
     */
    [TestFixture]
    public class TestMinusZeroResult
    {
        private static double MINUS_ZERO = -0.0;

        // convenient access to namepace
        //no use private static EvalInstances EI = null;
        [Test]
        public void TestSimpleOperators()
        {

            // unary plus is a no-op
            CheckEval(MINUS_ZERO, UnaryPlusEval.instance, MINUS_ZERO);

            // most simple operators convert -0.0 to +0.0
            CheckEval(0.0, EvalInstances.UnaryMinus, 0.0);
            CheckEval(0.0, EvalInstances.Percent, MINUS_ZERO);
            CheckEval(0.0, EvalInstances.Multiply, MINUS_ZERO, 1.0);
            CheckEval(0.0, EvalInstances.Divide, MINUS_ZERO, 1.0);
            CheckEval(0.0, EvalInstances.Power, MINUS_ZERO, 1.0);

            // but SubtractEval does not convert -0.0, so '-' and '+' work like java
            CheckEval(MINUS_ZERO, EvalInstances.Subtract, MINUS_ZERO, 0.0); // this is the main point of bug 47198
            CheckEval(0.0, EvalInstances.Add, MINUS_ZERO, 0.0);
        }

        /**
         * These results are hard to see in Excel (since -0.0 is usually Converted to +0.0 before it
         * Gets to the comparison operator)
         */
        [Test]
        [Ignore("no representation for -0.0 in c#, this test always failed.")]
        public void TestComparisonOperators()
        {
            CheckEval(false, EvalInstances.Equal, 0.0, MINUS_ZERO);
            CheckEval(true, EvalInstances.GreaterThan, 0.0, MINUS_ZERO);
            CheckEval(true, EvalInstances.LessThan, MINUS_ZERO, 0.0);
            
            //CheckEval(true, EvalInstances.Equal, 0.0, MINUS_ZERO);
            //CheckEval(false, EvalInstances.GreaterThan, 0.0, MINUS_ZERO);
            //CheckEval(false, EvalInstances.LessThan, MINUS_ZERO, 0.0);
        }
        [Test]
        public void TestTextRendering()
        {
            ConfirmTextRendering("-0", MINUS_ZERO);
            // sub-normal negative numbers also display as '-0'
            ConfirmTextRendering("-0", BitConverter.Int64BitsToDouble(unchecked((long)0x8000100020003000L)));
        }

        /**
         * Uses {@link ConcatEval} to force number-to-text conversion
         */
        private static void ConfirmTextRendering(String expRendering, double d)
        {
            ValueEval[] args = { StringEval.EMPTY_INSTANCE, new NumberEval(d), };
            StringEval se = (StringEval)EvalInstances.Concat.Evaluate(args, -1, (short)-1);
            String result = se.StringValue;
            Assert.AreEqual(expRendering, result);
        }

        private static void CheckEval(double expectedResult, Function instance, params double[] dArgs)
        {
            NumberEval result = (NumberEval)Evaluate(instance, dArgs);
            assertDouble(expectedResult, result.NumberValue);
        }
        private static void CheckEval(bool expectedResult, Function instance, params double[] dArgs)
        {
            BoolEval result = (BoolEval)Evaluate(instance, dArgs);
            Assert.AreEqual(expectedResult, result.BooleanValue);
        }
        private static ValueEval Evaluate(Function instance, params double[] dArgs)
        {
            ValueEval[] evalArgs;
            evalArgs = new ValueEval[dArgs.Length];
            for (int i = 0; i < evalArgs.Length; i++)
            {
                evalArgs[i] = new NumberEval(dArgs[i]);
            }
            ValueEval r = instance.Evaluate(evalArgs, -1, (short)-1);
            return r;
        }

        /**
         * Not really a POI Test - just Shows similar behaviour of '-0.0' in Java.
         */
        [Test]
        public void TestJava()
        {

            Assert.AreEqual(unchecked((long)0x8000000000000000L), BitConverter.DoubleToInt64Bits(MINUS_ZERO));

            // The simple operators consider all zeros to be the same
            //Assert.IsTrue(MINUS_ZERO == MINUS_ZERO);
            Assert.IsTrue(MINUS_ZERO == +0.0);
            Assert.IsFalse(MINUS_ZERO < +0.0);

            // Double.Compare() considers them different
            //Assert.IsTrue((MINUS_ZERO - +0.0) < 0);

            // multiplying zero by any negative quantity yields minus zero
            assertDouble(MINUS_ZERO, 0.0 * -1);
            assertDouble(MINUS_ZERO, 0.0 * -1e300);
            assertDouble(MINUS_ZERO, 0.0 * -1e-300);

            // minus zero can be produced as a result of underflow
            assertDouble(MINUS_ZERO, -1e-300 / 1e100);

            // multiplying or dividing minus zero by a positive quantity yields minus zero
            assertDouble(MINUS_ZERO, MINUS_ZERO * 1.0);
            assertDouble(MINUS_ZERO, MINUS_ZERO / 1.0);

            // subtracting positive zero gives minus zero
            assertDouble(MINUS_ZERO, MINUS_ZERO - 0.0);
            // BUT Adding positive zero gives positive zero
            assertDouble(0.0, MINUS_ZERO + 0.0);  // <<----
        }

        /**
         * Just so there is no ambiguity.  The two double values have to be exactly equal
         */
        private static void assertDouble(double a, double b)
        {
            long bitsA = BitConverter.DoubleToInt64Bits(a);
            long bitsB = BitConverter.DoubleToInt64Bits(b);
            if (bitsA != bitsB)
            {
                throw new ComparisonFailure("value different to expected",
                        new String(HexDump.LongToHex(bitsA)), new String(HexDump.LongToHex(bitsB)));
            }
        }
    }

}