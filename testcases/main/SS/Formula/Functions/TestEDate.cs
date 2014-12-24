using NPOI.SS.Formula.Atp;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestEDate
    {
        [Test]
        public void TestEDateProperValues()
        {
            EDate eDate = (EDate)EDate.Instance;
            NumberEval result = (NumberEval)eDate.Evaluate(new ValueEval[] { new NumberEval(1000), new NumberEval(0) }, null);
            Assert.AreEqual(1000d, result.NumberValue);
        }
        [Test]
        public void TestEDateInvalidValues()
        {
            EDate eDate = (EDate)EDate.Instance;
            ErrorEval result = (ErrorEval)eDate.Evaluate(new ValueEval[] { new NumberEval(1000) }, null);
            Assert.AreEqual(ErrorConstants.ERROR_VALUE, result.ErrorCode);
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
            Assert.AreEqual(a.Year, b.Year);
            Assert.AreEqual(a.Month, b.Month);
            Assert.AreEqual(a.Day, b.Day);
            Assert.AreEqual(a.Hour, b.Hour);
            Assert.AreEqual(a.Minute, b.Minute);
            Assert.AreEqual(a.Second, b.Second);
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
    }
}
