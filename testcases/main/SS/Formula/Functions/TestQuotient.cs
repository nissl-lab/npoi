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
    using System;

    /**
     * Tests for {@link Quotient}
     *
     * @author cedric dot walter @ gmail dot com
     */
    [TestFixture]
    public class TestQuotient
    {
        private static ValueEval invokeValue(String numerator, String denominator)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(numerator), new StringEval(denominator) };
            return new Quotient().Evaluate(args, -1, -1);
        }

        private static void ConfirmValue(String msg, String numerator, String denominator, String expected)
        {
            ValueEval result = invokeValue(numerator, denominator);
            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual(expected, ((NumberEval)result).StringValue, msg);
        }

        private static void ConfirmValueError(String msg, String numerator, String denominator, ErrorEval numError)
        {
            ValueEval result = invokeValue(numerator, denominator);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(numError, result, msg);
        }
        [Test]
        public void TestBasic()
        {
            ConfirmValue("int portion of 5/2 (2)", "5", "2", "2");
            ConfirmValue("int portion of 4.5/3.1 (1)", "4.5", "3.1", "1");

            ConfirmValue("int portion of -10/3 (-3)", "-10", "3", "-3");
            ConfirmValue("int portion of -5.5/2 (-2)", "-5.5", "2", "-2");

            ConfirmValue("int portion of Pi/Avogadro (0)", "3.14159", "6.02214179E+23", "0");
        }
        [Test]
        public void TestErrors()
        {
            ConfirmValueError("numerator is nonnumeric", "ABCD", "", ErrorEval.VALUE_INVALID);
            ConfirmValueError("denominator is nonnumeric", "", "ABCD", ErrorEval.VALUE_INVALID);

            ConfirmValueError("dividing by zero", "3.14159", "0", ErrorEval.DIV_ZERO);
        }
    }
}