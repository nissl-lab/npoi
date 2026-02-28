using System;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NUnit.Framework;using NUnit.Framework.Legacy;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestEDate
    {
        [Test]
        public void TestEDateProperValues()
        {
            // verify some border-case combinations of startDate and month-increase
            checkValue(1000, 0, 1000d);
            checkValue(1, 0, 1d);
            checkValue(0, 1, 31d);
            checkValue(1, 1, 32d);
            checkValue(0, 0, /* BAD_DATE! */ -1.0d);
            checkValue(0, -2, /* BAD_DATE! */ -1.0d);
            checkValue(0, -3, /* BAD_DATE! */ -1.0d);
            checkValue(49104, 0, 49104d);
            checkValue(49104, 1, 49134d);
        }

        private void checkValue(int startDate, int monthInc, double expectedResult)
        {
            EDate eDate = new EDate();
            NumberEval result = (NumberEval)eDate.Evaluate(new ValueEval[] { new NumberEval(startDate), new NumberEval(monthInc) }, null);
            ClassicAssert.AreEqual(expectedResult, result.NumberValue);
        }
        [Test]
        public void TestEDateInvalidValues()
        {
            EDate eDate = (EDate)EDate.Instance;
            ErrorEval result = (ErrorEval)eDate.Evaluate(new ValueEval[] { new NumberEval(1000) }, null);
            ClassicAssert.AreEqual(FormulaError.VALUE.Code, result.ErrorCode);
        }
        [Test]
        public void TestEDateIncrease()
        {
            EDate eDate = (EDate)EDate.Instance;
            DateTime startDate = DateTime.Now;
            int offset = 2;
            NumberEval result = (NumberEval)eDate.Evaluate(new ValueEval[] { new NumberEval(DateUtil.GetExcelDate(startDate)), new NumberEval(offset) }, null);
            DateTime resultDate = DateUtil.GetJavaDate(result.NumberValue);
            CompareDateTimes(resultDate, startDate.AddMonths(offset));

        }
        private void CompareDateTimes(DateTime a, DateTime b)
        {
            ClassicAssert.AreEqual(a.Year, b.Year);
            ClassicAssert.AreEqual(a.Month, b.Month);
            ClassicAssert.AreEqual(a.Day, b.Day);
            ClassicAssert.AreEqual(a.Hour, b.Hour);
            ClassicAssert.AreEqual(a.Minute, b.Minute);
            ClassicAssert.AreEqual(a.Second, b.Second, delta: 1); // can shift during tests
        }
        [Test]
        public void TestEDateDecrease()
        {
            EDate eDate = (EDate)EDate.Instance;
            DateTime startDate = DateTime.Now;
            int offset = -2;
            NumberEval result = (NumberEval)eDate.Evaluate(new ValueEval[] { new NumberEval(DateUtil.GetExcelDate(startDate)), new NumberEval(offset) }, null);
            DateTime resultDate = DateUtil.GetJavaDate(result.NumberValue);

            CompareDateTimes(resultDate, startDate.AddMonths(offset));
            
        }

        [Test]
        public void TestBug56688()
        {
            EDate eDate = new EDate();
            NumberEval result = (NumberEval)eDate.Evaluate(new ValueEval[] { new NumberEval(1000), new RefEvalImplementation(new NumberEval(0)) }, null);
            ClassicAssert.AreEqual(1000d, result.NumberValue);
        }

        [Test]
        public void TestRefEvalStartDate()
        {
            EDate eDate = new EDate();
            NumberEval result = (NumberEval)eDate.Evaluate(new ValueEval[] { new RefEvalImplementation(new NumberEval(1000)), new NumberEval(0) }, null);
            ClassicAssert.AreEqual(1000d, result.NumberValue);
        }
        private class ValueEval1 : ValueEval
        {

        }
        [Test]
        public void TestEDateInvalidValueEval()
        {
            ValueEval Evaluate = new EDate().Evaluate(new ValueEval[] { new ValueEval1() { }, new NumberEval(0) }, null);
            ClassicAssert.IsTrue(Evaluate is ErrorEval);
            ClassicAssert.AreEqual(ErrorEval.VALUE_INVALID, Evaluate);
        }

        [Test]
        public void TestEDateBlankValueEval()
        {
            NumberEval Evaluate = (NumberEval)new EDate().Evaluate(new ValueEval[] { BlankEval.instance, new NumberEval(0) }, null);
            ClassicAssert.AreEqual(-1.0d, Evaluate.NumberValue);
        }

        [Test]
        public void TestEDateBlankRefValueEval()
        {
            EDate eDate = new EDate();
            NumberEval result = (NumberEval)eDate.Evaluate(new ValueEval[] { new RefEvalImplementation(BlankEval.instance), new NumberEval(0) }, null);
            ClassicAssert.AreEqual(-1.0d, result.NumberValue, "0 startDate triggers BAD_DATE currently, thus -1.0!");

            result = (NumberEval)eDate.Evaluate(new ValueEval[] { new NumberEval(1), new RefEvalImplementation(BlankEval.instance) }, null);
            ClassicAssert.AreEqual(1.0d, result.NumberValue, "Blank is handled as 0 otherwise");
        }

    }
}
