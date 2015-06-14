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
     * Tests for {@link Complex}
     *
     * @author cedric dot walter @ gmail dot com
     */
    [TestFixture]
    public class TestComplex
    {
        private static ValueEval invokeValue(String real_num, String i_num, String suffix)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(real_num), new StringEval(i_num), new StringEval(suffix) };
            return new Complex().Evaluate(args, -1, -1);
        }

        private static void ConfirmValue(String msg, String real_num, String i_num, String suffix, String expected)
        {
            ValueEval result = invokeValue(real_num, i_num, suffix);
            Assert.AreEqual(typeof(StringEval), result.GetType());
            Assert.AreEqual( expected, ((StringEval)result).StringValue, msg);
        }

        private static void ConfirmValueError(String msg, String real_num, String i_num, String suffix, ErrorEval numError)
        {
            ValueEval result = invokeValue(real_num, i_num, suffix);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual( numError, result, msg);
        }
        [Test]
        public void TestBasic()
        {
            ConfirmValue("Complex number with 3 and 4 as the real and imaginary coefficients (3 + 4i)", "3", "4", "", "3+4i");
            ConfirmValue("Complex number with 3 and 4 as the real and imaginary coefficients, and j as the suffix (3 + 4j)", "3", "4", "j", "3+4j");

            ConfirmValue("Complex number with 0 and 1 as the real and imaginary coefficients (i)", "0", "1", "", "i");
            ConfirmValue("Complex number with 1 and 0 as the real and imaginary coefficients (1)", "1", "0", "", "1");

            ConfirmValue("Complex number with 2 and 3 as the real and imaginary coefficients (2 + 3i)", "2", "3", "", "2+3i");
            ConfirmValue("Complex number with -2 and -3 as the real and imaginary coefficients (-2-3i)", "-2", "-3", "", "-2-3i");

            ConfirmValue("Complex number with -2 and -3 as the real and imaginary coefficients (-0.5-3.2i)", "-0.5", "-3.2", "", "-0.5-3.2i");
        }
        [Test]
        public void TestErrors()
        {
            ConfirmValueError("argument is nonnumeric", "ABCD", "", "", ErrorEval.VALUE_INVALID);
            ConfirmValueError("argument is nonnumeric", "1", "ABCD", "", ErrorEval.VALUE_INVALID);
            ConfirmValueError("f suffix is neither \"i\" nor \"j\"", "1", "1", "k", ErrorEval.VALUE_INVALID);

            ConfirmValueError("never use \"I\" ", "1", "1", "I", ErrorEval.VALUE_INVALID);
            ConfirmValueError("never use \"J\" ", "1", "1", "J", ErrorEval.VALUE_INVALID);
        }
    }
}