using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
	[TestFixture]
	public class TestDateValue
	{
		private ValueEval invokeDateValue(ValueEval text)
		{
			return new DateValue().Evaluate(0, 0, text);
		}

		private void confirmDateValue(ValueEval text, double expected)
		{
			ValueEval result = invokeDateValue(text);
			Assert.IsTrue(result is NumberEval);
			Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0.0001);
		}

		private void confirmDateValue(ValueEval text, BlankEval expected)
		{
			ValueEval result = invokeDateValue(text);
			Assert.IsTrue(result is BlankEval);
		}

		private void confirmDateValue(ValueEval text, ErrorEval expected)
		{
			ValueEval result = invokeDateValue(text);
			Assert.IsTrue(result is ErrorEval);
			Assert.AreEqual(expected.ErrorCode, ((ErrorEval)result).ErrorCode);
		}
		[Test]
		public void TestDateValues()
		{
			confirmDateValue(new StringEval("2020-02-01"), 43862);
			confirmDateValue(new StringEval("01-02-2020"), 43862);
			confirmDateValue(new StringEval("2020-FEB-01"), 43862);
			confirmDateValue(new StringEval("2020-Feb-01"), 43862);
			confirmDateValue(new StringEval("2020-FEBRUARY-01"), 43862);
			//confirmDateValue(new StringEval("FEB-01"), 43862);
			confirmDateValue(new StringEval("2/1/2020"), 43862);
			//confirmDateValue(new StringEval("2/1"), 43862);
			confirmDateValue(new StringEval("2020/2/1"), 43862);
			confirmDateValue(new StringEval("2020/FEB/1"), 43862);
			confirmDateValue(new StringEval("FEB/1/2020"), 43862);
			confirmDateValue(new StringEval("2020/02/01"), 43862);

			confirmDateValue(new StringEval(""), BlankEval.instance);
			confirmDateValue(BlankEval.instance, BlankEval.instance);

			confirmDateValue(new StringEval("non-date text"), ErrorEval.VALUE_INVALID);
		}
	}
}
