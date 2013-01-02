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
    using NPOI.SS.Formula.Functions;
    using NUnit.Framework;
    using System;

    /**
     * Tests for Excel function MID()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestMid
    {


        private static ValueEval invokeMid(ValueEval text, ValueEval startPos, ValueEval numChars)
        {
            ValueEval[] args = new ValueEval[] { text, startPos, numChars, };
            return TextFunction.MID.Evaluate(args, -1, (short)-1);
        }

        private void ConfirmMid(ValueEval text, ValueEval startPos, ValueEval numChars, String expected)
        {
            ValueEval result = invokeMid(text, startPos, numChars);
            Assert.AreEqual(typeof(StringEval), result.GetType());
            Assert.AreEqual(expected, ((StringEval)result).StringValue);
        }

        private void ConfirmMid(ValueEval text, ValueEval startPos, ValueEval numChars, ErrorEval expectedError)
        {
            ValueEval result = invokeMid(text, startPos, numChars);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(expectedError.ErrorCode, ((ErrorEval)result).ErrorCode);
        }
        [Test]
        public void TestBasic()
        {

            ConfirmMid(new StringEval("galactic"), new NumberEval(3), new NumberEval(4), "lact");
        }

        /**
         * Valid cases where args are not precisely (string, int, int) but can be Resolved OK.
         */
        [Test]
        public void TestUnusualArgs()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("en-US");

            // startPos with fractional digits
            ConfirmMid(new StringEval("galactic"), new NumberEval(3.1), new NumberEval(4), "lact");

            // string startPos
            ConfirmMid(new StringEval("galactic"), new StringEval("3"), new NumberEval(4), "lact");

            // text (first) arg type is number, other args are strings with fractional digits
            ConfirmMid(new NumberEval(123456), new StringEval("3.1"), new StringEval("2.9"), "34");

            // startPos is 1x1 area ref, numChars is cell ref
            AreaEval aeStart = EvalFactory.CreateAreaEval("A1:A1", new ValueEval[] { new NumberEval(2), });
            RefEval reNumChars = EvalFactory.CreateRefEval("B1", new NumberEval(3));
            ConfirmMid(new StringEval("galactic"), aeStart, reNumChars, "ala");

            ConfirmMid(new StringEval("galactic"), new NumberEval(3.1), BlankEval.instance, "");

            ConfirmMid(new StringEval("galactic"), new NumberEval(3), BoolEval.FALSE, "");
            ConfirmMid(new StringEval("galactic"), new NumberEval(3), BoolEval.TRUE, "l");
            ConfirmMid(BlankEval.instance, new NumberEval(3), BoolEval.TRUE, "");

        }

        /**
         * Extreme values for startPos and numChars
         */
        [Test]
        public void TestExtremes()
        {
            ConfirmMid(new StringEval("galactic"), new NumberEval(4), new NumberEval(400), "actic");

            ConfirmMid(new StringEval("galactic"), new NumberEval(30), new NumberEval(4), "");
            ConfirmMid(new StringEval("galactic"), new NumberEval(3), new NumberEval(0), "");
        }

        /**
         * All sorts of ways to make MID return defined errors.
         */
        [Test]
        public void TestErrors()
        {
            ConfirmMid(ErrorEval.NAME_INVALID, new NumberEval(3), new NumberEval(4), ErrorEval.NAME_INVALID);
            ConfirmMid(new StringEval("galactic"), ErrorEval.NAME_INVALID, new NumberEval(4), ErrorEval.NAME_INVALID);
            ConfirmMid(new StringEval("galactic"), new NumberEval(3), ErrorEval.NAME_INVALID, ErrorEval.NAME_INVALID);
            ConfirmMid(new StringEval("galactic"), ErrorEval.DIV_ZERO, ErrorEval.NAME_INVALID, ErrorEval.DIV_ZERO);

            ConfirmMid(new StringEval("galactic"), BlankEval.instance, new NumberEval(3.1), ErrorEval.VALUE_INVALID);

            ConfirmMid(new StringEval("galactic"), new NumberEval(0), new NumberEval(4), ErrorEval.VALUE_INVALID);
            ConfirmMid(new StringEval("galactic"), new NumberEval(1), new NumberEval(-1), ErrorEval.VALUE_INVALID);
        }
    }

}