using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestCode
    {
        private static ValueEval invokeValue(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), };
            return new Code().Evaluate(args, -1, -1);
        }

        private static void confirmValue(String msg, String number1, String expected)
        {
            ValueEval result = invokeValue(number1);
            Assert.AreEqual(typeof(StringEval), result.GetType());
            Assert.AreEqual(expected, ((StringEval)result).StringValue, msg);
        }

        private static void confirmValueError(String msg, String number1, ErrorEval numError)
        {
            ValueEval result = invokeValue(number1);
            Assert.AreEqual(typeof(ErrorEval), result.GetType());
            Assert.AreEqual(numError, result, msg);
        }

        [Test]
        public void TestBasic()
        {
            confirmValue("Displays the numeric code for A (65)", "A", "65");
            confirmValue("Displays the numeric code for the first character in text ABCDEFGHI (65)", "ABCDEFGHI", "65");

            confirmValue("Displays the numeric code for ! (33)", "!", "33");
        }
        [Test]
        public void TestErrors()
        {
            confirmValueError("Empty text", "", ErrorEval.VALUE_INVALID);
        }
    }
}
