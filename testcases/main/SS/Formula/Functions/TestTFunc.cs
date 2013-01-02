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
    using System;
    using NPOI.SS.Formula.Functions;


    /**
     * Test cases for Excel function T()
     *
     * @author Josh Micich
     */
    [TestFixture]
    public class TestTFunc
    {

        /**
         * @return the result of calling function T() with the specified argument
         */
        
        private static ValueEval invokeT(ValueEval arg)
        {
            ValueEval[] args = { arg, };
            ValueEval result = new T().Evaluate(args, -1, (short)-1);
            Assert.IsNotNull(result, "result may never be null");
            return result;
        }
        /**
         * Simulates call: T(A1)
         * where cell A1 has the specified innerValue
         */
        private ValueEval invokeTWithReference(ValueEval innerValue)
        {
            ValueEval arg = EvalFactory.CreateRefEval("$B$2", innerValue);
            return invokeT(arg);
        }

        private static void ConfirmText(String text)
        {
            ValueEval arg = new StringEval(text);
            ValueEval eval = invokeT(arg);
            StringEval se = (StringEval)eval;
            Assert.AreEqual(text, se.StringValue);
        }
        [Test]
        public void TestTextValues()
        {

            ConfirmText("abc");
            ConfirmText("");
            ConfirmText(" ");
            ConfirmText("~");
            ConfirmText("123");
            ConfirmText("TRUE");
        }

        private static void ConfirmError(ValueEval arg)
        {
            ValueEval eval = invokeT(arg);
            Assert.IsTrue(arg == eval);
        }
        [Test]
        public void TestErrorValues()
        {

            ConfirmError(ErrorEval.VALUE_INVALID);
            ConfirmError(ErrorEval.NA);
            ConfirmError(ErrorEval.REF_INVALID);
        }

        private static void ConfirmString(ValueEval eval, String expected)
        {
            Assert.IsTrue(eval is StringEval);
            Assert.AreEqual(expected, ((StringEval)eval).StringValue);
        }

        private static void ConfirmOther(ValueEval arg)
        {
            ValueEval eval = invokeT(arg);
            ConfirmString(eval, "");
        }
        [Test]
        public void TestOtherValues()
        {
            ConfirmOther(new NumberEval(2));
            ConfirmOther(BoolEval.FALSE);
            ConfirmOther(BlankEval.instance);  // can this particular case be verified?
        }
        [Test]
        public void TestRefValues()
        {
            ValueEval eval;

            eval = invokeTWithReference(new StringEval("def"));
            ConfirmString(eval, "def");
            eval = invokeTWithReference(new StringEval(" "));
            ConfirmString(eval, " ");

            eval = invokeTWithReference(new NumberEval(2));
            ConfirmString(eval, "");
            eval = invokeTWithReference(BoolEval.TRUE);
            ConfirmString(eval, "");

            eval = invokeTWithReference(ErrorEval.NAME_INVALID);
            Assert.IsTrue(eval == ErrorEval.NAME_INVALID);
        }
        [Test]
        public void TestAreaArg()
        {
            ValueEval[] areaValues = new ValueEval[] {
			new StringEval("abc"), new StringEval("def"),
			new StringEval("ghi"), new StringEval("jkl"),
		};
            AreaEval ae = EvalFactory.CreateAreaEval("C10:D11", areaValues);

            ValueEval ve;
            ve = invokeT(ae);
            ConfirmString(ve, "abc");

            areaValues[0] = new NumberEval(5.0);
            ve = invokeT(ae);
            ConfirmString(ve, "");
        }
    }

}