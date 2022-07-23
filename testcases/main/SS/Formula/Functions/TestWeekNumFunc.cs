using NPOI.SS.Formula;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestWeekNumFunc
    {
        private static double TOLERANCE = 0.001;
        private static OperationEvaluationContext DEFAULT_CONTEXT =
            new OperationEvaluationContext(null, null, 0, 1, 0, null);

        [Test]
        public void TestEvaluate()
        {
            var localDate = DateTime.Parse("2012-03-09");
            double date = DateUtil.GetExcelDate(localDate);
            assertEvaluateEquals(10.0, date);
            assertEvaluateEquals(10.0, date, 1);
            assertEvaluateEquals(11.0, date, 2);
            assertEvaluateEquals(11.0, date, 11);
            assertEvaluateEquals(11.0, date, 12);
            assertEvaluateEquals(11.0, date, 13);
            assertEvaluateEquals(11.0, date, 14);
            assertEvaluateEquals(11.0, date, 15);
            assertEvaluateEquals(10.0, date, 16);
            assertEvaluateEquals(10.0, date, 17);
            assertEvaluateEquals(10.0, date, 21);
        }
        [Test]
        public void TestEvaluateInvalid()
        {
            assertEvaluateEquals("no args", ErrorEval.VALUE_INVALID);
            assertEvaluateEquals("too many args", ErrorEval.VALUE_INVALID, new NumberEval(1.0), new NumberEval(1.0), new NumberEval(1.0));
            assertEvaluateEquals("negative date", ErrorEval.NUM_ERROR, new NumberEval(-1.0));
            assertEvaluateEquals("cannot coerce serial_number to number", ErrorEval.VALUE_INVALID, new StringEval(""));
            assertEvaluateEquals("cannot coerce return_type to number", ErrorEval.NUM_ERROR, new NumberEval(1.0), new StringEval(""));
            assertEvaluateEquals("return_type is blank", ErrorEval.NUM_ERROR, new StringEval("2"), BlankEval.instance);
            assertEvaluateEquals("invalid return_type", ErrorEval.NUM_ERROR, new NumberEval(1.0), new NumberEval(18.0));
        }

        // for testing invalid invocations
        private void assertEvaluateEquals(String message, ErrorEval expected, params ValueEval[] args)
        {
            String formula = "WEEKNUM(" + StringUtil.Join(args, ", ") + ")";
            ValueEval result = WeekNum.instance.Evaluate(args, DEFAULT_CONTEXT);
            Assert.AreEqual(expected, result, formula + ": " + message);
        }

        private void assertEvaluateEquals(double expected, double dateValue)
        {
            String formula = "WEEKNUM(" + dateValue + ")";
            ValueEval[] args = new ValueEval[] { new NumberEval(dateValue) };
            ValueEval result = WeekNum.instance.Evaluate(args, DEFAULT_CONTEXT);
            if (result is NumberEval) {
                Assert.AreEqual(expected, ((NumberEval)result).NumberValue, TOLERANCE, formula);
            } else
            {
                Assert.Fail("unexpected eval result " + result);
            }
        }

        private void assertEvaluateEquals(double expected, double dateValue, double return_type)
        {
            String formula = "WEEKNUM(" + dateValue + ", " + return_type + ")";
            ValueEval[] args = new ValueEval[] { new NumberEval(dateValue), new NumberEval(return_type) };
            ValueEval result = WeekNum.instance.Evaluate(args, DEFAULT_CONTEXT);
            if (result is NumberEval) {
                Assert.AreEqual(expected, ((NumberEval)result).NumberValue, TOLERANCE, formula);
            } else
            {
                Assert.Fail("unexpected eval result " + result);
            }
        }
    }
}
