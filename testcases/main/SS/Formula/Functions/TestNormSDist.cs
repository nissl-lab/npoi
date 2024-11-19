using NPOI.SS.Formula;
using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestNormSDist
    {
        private static OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);
        [Test]
        public void TestBasic()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            confirmValue("1.333333", 0.908788726);
        }

        [Test]
        public void TestInvalid()
        {
            confirmInvalidError("A1");
        }
        [Test]
        public void TestMicrosoftExample1()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = row.CreateCell(0);
            SS.Util.Utils.AssertDouble(fe, cell, "NORMSDIST(1.333333)", 0.908788726, 0.000001);
            SS.Util.Utils.AssertDouble(fe, cell, "NORM.S.DIST(1.333333)", 0.908788726, 0.000001);
        }
        private static ValueEval invokeValue(String number1)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1) };
            return NormSDist.instance.Evaluate(args, ec);
        }

        private static void confirmValue(String number1, double expected)
        {
            ValueEval result = invokeValue(number1);
            Assert.IsTrue(result is NumberEval);
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0.0000001);
        }

        private static void confirmInvalidError(String number1)
        {
            ValueEval result = invokeValue(number1);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }
    }
}
