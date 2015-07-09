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
     * Tests for {@link Hex2Dec}
     *
     * @author cedric dot walter @ gmail dot com
     */
    [TestFixture]
    public class TestHex2Dec
    {

        private static ValueEval invokeValue(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1) };
            return new Hex2Dec().Evaluate(args, -1, -1);
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
            ConfirmValue("Converts octal 'A5' to decimal (165)", "A5", "165");
            ConfirmValue("Converts octal FFFFFFFF5B to decimal (-165)", "FFFFFFFF5B", "-165");
            ConfirmValue("Converts octal 3DA408B9 to decimal (-165)", "3DA408B9", "1034160313");
        }

        [Test]
        public void TestErrors()
        {
            ConfirmValueError("not a valid octal number", "GGGGGGG", ErrorEval.NUM_ERROR);
            ConfirmValueError("not a valid octal number", "3.14159", ErrorEval.NUM_ERROR);
        }
    }

}