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
     * Tests for {@link Dec2Hex}
     *
     * @author cedric dot walter @ gmail dot com
     */
    [TestFixture]
    public class TestDec2Hex
    {

        private static ValueEval invokeValue(String number1, String number2)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), new StringEval(number2), };
            return new Dec2Hex().Evaluate(args, -1, -1);
        }

        private static ValueEval invokeValue(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), };
            return new Dec2Hex().Evaluate(args, -1, -1);
        }

        private static void ConfirmValue(String msg, String number1, String number2, String expected)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.AreEqual(typeof(StringEval), result.GetType());
            Assert.AreEqual(expected, ((StringEval)result).StringValue, msg);
        }

        private static void ConfirmValue(String msg, String number1, String expected)
        {
            ValueEval result = invokeValue(number1);
            Assert.AreEqual(typeof(StringEval), result.GetType());
            Assert.AreEqual(expected, ((StringEval)result).StringValue, msg);
        }

        private static void ConfirmValueError(String msg, String number1, String number2, ErrorEval numError)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(numError, result, msg);
        }
        [Test]
        public void TestBasic()
        {
            ConfirmValue("Converts decimal 100 to hexadecimal with 0 characters (64)", "100", "0", "64");
            ConfirmValue("Converts decimal 100 to hexadecimal with 4 characters (0064)", "100", "4", "0064");
            ConfirmValue("Converts decimal 100 to hexadecimal with 5 characters (0064)", "100", "5", "00064");
            ConfirmValue("Converts decimal 100 to hexadecimal with 10 (default) characters", "100", "10", "0000000064");
            ConfirmValue("If argument places Contains a decimal value, dec2hex ignores the numbers to the right side of the decimal point.", "100", "10.0", "0000000064");

            ConfirmValue("Converts decimal -54 to hexadecimal, 0 is ignored", "-54", "0", "00000FFFCA");
            ConfirmValue("Converts decimal -54 to hexadecimal, 2 is ignored", "-54", "2", "00000FFFCA");
            ConfirmValue("places is optionnal", "-54", "00000FFFCA");
        }
        [Test]
        public void TestErrors()
        {
            ConfirmValueError("Out of range min number", "-549755813889", "0", ErrorEval.NUM_ERROR);
            ConfirmValueError("Out of range max number", "549755813888", "0", ErrorEval.NUM_ERROR);

            ConfirmValueError("negative places not allowed", "549755813888", "-10", ErrorEval.NUM_ERROR);
            ConfirmValueError("non number places not allowed", "ABCDEF", "0", ErrorEval.VALUE_INVALID);
        }
    }

}