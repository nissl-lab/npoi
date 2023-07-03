using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NUnit.Framework;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestStandardize
    {
        private static OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);
        [Test]
        public void TestBasic()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            confirmValue("42", "40", "1.5", 1.33333333);
        }

        [Test]
        public void TestInvalid()
        {
            confirmInvalidError("A1", "B2", "C2");
        }

        [Test]
        public void TestNumError()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            confirmNumError("42", "40", "0");
            confirmNumError("42", "40", "-0.1");
        }
        //https://support.microsoft.com/en-us/office/standardize-function-81d66554-2d54-40ec-ba83-6437108ee775
        [Test]
        public void TestMicrosoftExample1()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, "Data", "Description");
            SS.Util.Utils.AddRow(sheet, 1, 42, "Value to normalize");
            SS.Util.Utils.AddRow(sheet, 2, 40, "Arithmetic mean of the distribution");
            SS.Util.Utils.AddRow(sheet, 3, 1.5, "Standard deviation of the distribution");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            SS.Util.Utils.AssertDouble(fe, cell, "STANDARDIZE(A2,A3,A4)", 1.33333333, 0.000001);
        }
        private static ValueEval invokeValue(String number1, String number2, String number3)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), new StringEval(number2), new StringEval(number3) };
            return Standardize.instance.Evaluate(args, ec);
        }

        private static void confirmValue(String number1, String number2, String number3, double expected)
        {
            ValueEval result = invokeValue(number1, number2, number3);
            Assert.IsTrue(result is NumberEval);
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0.0000001);
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