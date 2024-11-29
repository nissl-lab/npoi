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
    [TestFixture]
    public class TestBesselJ
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

        //https://support.microsoft.com/en-us/office/besselj-function-839cb181-48de-408b-9d80-bd02982d94f7
        [Test]
        public void TestMicrosoftExample1()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            IRow row = sheet.CreateRow(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = row.CreateCell(0);
            //this tolerance is too high but commons-math3 and excel don't match up as closely as we'd like
            double tolerance = 0.000001;
            SS.Util.Utils.AssertDouble(fe, cell, "BESSELJ(1.9, 2)", 0.329925829, tolerance);
            SS.Util.Utils.AssertDouble(fe, cell, "BESSELJ(1.9, 2.5)", 0.329925829, tolerance);
        }

        private static ValueEval invokeValue(String number1, String number2)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), new StringEval(number2) };
            return BesselJ.instance.Evaluate(args, ec);
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
    }
}
