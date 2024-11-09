using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestDollarFr
    {
        private static OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);

        [Test]
        public void TestInvalid()
        {
            confirmInvalidError("A1", "B2");
        }

        [Test]
        public void TestNumError()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            confirmNumError("22.5", "-40");
        }

        [Test]
        public void TestDiv0()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            confirmDiv0("22.5", "0");
            confirmDiv0("22.5", "0.9");
            confirmDiv0("22.5", "-0.9");
        }

        //https://support.microsoft.com/en-us/office/dollarde-function-db85aab0-1677-428a-9dfd-a38476693427
        [Test]
        public void TestMicrosoftExample1()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            HSSFWorkbook wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet();
            var row = sheet.CreateRow(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            var cell = row.CreateCell(0);
            double tolerance = 0.000000000000001;
            SS.Util.Utils.AssertDouble(fe, cell, "DOLLARFR(1.125,16)", 1.02, tolerance);
            SS.Util.Utils.AssertDouble(fe, cell, "DOLLARFR(-1.125,16)", -1.02, tolerance);
            SS.Util.Utils.AssertDouble(fe, cell, "DOLLARFR(1.000125,16)", 1.00002, tolerance);
            SS.Util.Utils.AssertDouble(fe, cell, "DOLLARFR(1.125,32)", 1.04, tolerance);
        }

        private static ValueEval invokeValue(String number1, String number2)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), new StringEval(number2) };
            return DollarDe.instance.Evaluate(args, ec);
        }

        private static void confirmValue(String number1, String number2, double expected)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.IsTrue(result is NumberEval);
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0.00000000000001);
        }

        private static void confirmInvalidError(String number1, String number2)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        private static void confirmNumError(String number1, String number2)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorEval.NUM_ERROR, result);
        }

        private static void confirmDiv0(String number1, String number2)
        {
            ValueEval result = invokeValue(number1, number2);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorEval.DIV_ZERO, result);
        }
    }
}
