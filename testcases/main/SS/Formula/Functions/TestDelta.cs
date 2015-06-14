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

using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using NUnit.Framework;
namespace TestCases.SS.Formula.Functions
{
    /**
     * Tests for {@link NPOI.ss.formula.functions.Delta}
     *
     * @author cedric dot walter @ gmail dot com
     */
    [TestFixture]
    public class TestDelta
    {

        private static ValueEval invokeValue(String number1, String number2)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), new StringEval(number2), };
            return new Delta().Evaluate(args, -1, -1);
        }

        private static void ConfirmValue(String number1, String number2, double expected)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0.0);
        }

        private static void ConfirmValueError(String number1, String number2)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }
        [Test]
        public void TestBasic()
        {
            ConfirmValue("5", "4", 0); // Checks whether 5 Equals 4 (0)
            ConfirmValue("5", "5", 1); // Checks whether 5 Equals 5 (1)

            ConfirmValue("0.5", "0", 0); // Checks whether 0.5 Equals 0 (0)
            ConfirmValue("0.50", "0.5", 1);
            ConfirmValue("0.5000000000", "0.5", 1);
        }
        [Test]
        public void TestErrors()
        {
            ConfirmValueError("A1", "B2");
            ConfirmValueError("AAAA", "BBBB");
        }
    }

}