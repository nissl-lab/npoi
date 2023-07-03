using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    public class TestTDist
    {
        private static OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);
        [Test]
        public void TestBasic()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            confirmValue("5.968191467", "8", "1", 0.00016754180265310392);
            confirmValue("5.968191467", "8", "2", 0.00033508360530620784);
            confirmValue("5.968191467", "8.2", "2.2", 0.00033508360530620784);
            confirmValue("5.968191467", "8.9", "2.9", 0.00033508360530620784);
        }

        [Test]
        public void TestInvalid()
        {
            confirmInvalidError("A1", "B2", "C2");
            confirmInvalidError("5.968191467", "8", "C2");
            confirmInvalidError("5.968191467", "B2", "2");
            confirmInvalidError("A1", "8", "2");
        }

        [Test]
        public void TestNumError()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            confirmNumError("5.968191467", "8", "0");
            confirmNumError("-5.968191467", "8", "2");
        }

        //https://support.microsoft.com/en-us/office/tdist-function-630a7695-4021-4853-9468-4a1f9dcdd192
        [Test]
        public void testMicrosoftExample1()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, "Data", "Description");
            SS.Util.Utils.AddRow(sheet, 1, 1.959999998, "Value at which to evaluate the distribution");
            SS.Util.Utils.AddRow(sheet, 2, 60, "Degrees of freedom");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            SS.Util.Utils.AssertDouble(fe, cell, "TDIST(A2,A3,2)", 0.054644930, 0.01);
            SS.Util.Utils.AssertDouble(fe, cell, "TDIST(A2,A3,1)", 0.027322465, 0.01);

        }
        private static ValueEval invokeValue(String number1, String number2, String number3)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), new StringEval(number2), new StringEval(number3) };
            return TDist.instance.Evaluate(args, ec);
        }

        private static void confirmValue(String number1, String number2, String number3, double expected)
        {
            ValueEval result = invokeValue(number1, number2, number3);
            Assert.IsTrue(result is NumberEval);
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0.0);
        }

        private static void confirmInvalidError(String number1, String number2, String number3)
        {
            ValueEval result = invokeValue(number1, number2, number3);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        private static void confirmNumError(String number1, String number2, String number3)
        {
            ValueEval result = invokeValue(number1, number2, number3);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorEval.NUM_ERROR, result);
        }
    }

}
