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
     * Tests for {@link NPOI.SS.Formula.Functions.Oct2Dec}
     *
     * @author cedric dot walter @ gmail dot com
     */
    [TestFixture]
    public class TestOct2Dec
    {

        private static ValueEval invokeValue(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1) };
            return new Oct2Dec().Evaluate(args, -1, -1);
        }

        private static void ConfirmValue(String msg, String number1, String expected)
        {
            ValueEval result = invokeValue(number1);
            Assert.AreEqual(typeof(NumberEval), result.GetType());
            Assert.AreEqual(expected, ((NumberEval)result).StringValue, msg);
        }

        private static void ConfirmValueError(String msg, String number1, ErrorEval numError)
        {
            ValueEval result = invokeValue(number1);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(numError, result, msg);
        }

        [Test]
        public void TestBasic()
        {
            ConfirmValue("Converts octal '' to decimal (0)", "", "0");
            ConfirmValue("Converts octal 54 to decimal (44)", "54", "44");
            ConfirmValue("Converts octal 7777777533 to decimal (-165)", "7777777533", "-165");
            ConfirmValue("Converts octal 7000000000 to decimal (-134217728)", "7000000000", "-134217728");
            ConfirmValue("Converts octal 7776667533 to decimal (-299173)", "7776667533", "-299173");
        }

        [Test]
        public void TestErrors()
        {
            ConfirmValueError("not a valid octal number", "ABCDEFGH", ErrorEval.NUM_ERROR);
            ConfirmValueError("not a valid octal number", "99999999", ErrorEval.NUM_ERROR);
            ConfirmValueError("not a valid octal number", "3.14159", ErrorEval.NUM_ERROR);
        }
    }

}