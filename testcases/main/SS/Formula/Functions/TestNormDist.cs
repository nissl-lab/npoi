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
    public class TestNormDist
    {
        private static OperationEvaluationContext ec = new OperationEvaluationContext(null, null, 0, 0, 0, null);
        [Test]
        public void TestBasic()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            confirmValue("42", "40", "1.5", true, 0.908788780274132);
            confirmValue("42", "40", "1.5", false, 0.109340049783996);
        }
        [Test]
        public void TestInvalid()
        {
            confirmInvalidError("A1", "B2", "C2", false);
        }

        [Test]
        public void TestNumError()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            confirmNumError("42", "40", "0", false);
            confirmNumError("42", "40", "0", true);
            confirmNumError("42", "40", "-0.1", false);
            confirmNumError("42", "40", "-0.1", true);
        }

        //https://support.microsoft.com/en-us/office/normdist-function-126db625-c53e-4591-9a22-c9ff422d6d58
        //https://support.microsoft.com/en-us/office/norm-dist-function-edb1cc14-a21c-4e53-839d-8082074c9f8d
        [Test]
        public void TestMicrosoftExample1()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, "Data", "Description");
            SS.Util.Utils.AddRow(sheet, 1, 42, "Value for which you want the distribution");
            SS.Util.Utils.AddRow(sheet, 2, 40, "Arithmetic mean of the distribution");
            SS.Util.Utils.AddRow(sheet, 3, 1.5, "Standard deviation of the distribution");
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            SS.Util.Utils.AssertDouble(fe, cell, "NORMDIST(A2,A3,A4,TRUE)", 0.908788780274132, 0.00000000000001);
            SS.Util.Utils.AssertDouble(fe, cell, "NORM.DIST(A2,A3,A4,TRUE)", 0.908788780274132, 0.00000000000001);
        }

        private static ValueEval invokeValue(String number1, String number2, String number3, bool cumulative)
        {
            ValueEval[] args = new ValueEval[] { new StringEval(number1), new StringEval(number2), new StringEval(number3), BoolEval.ValueOf(cumulative) };
            return NormDist.instance.Evaluate(args, ec);
        }

        private static void confirmValue(String number1, String number2, String number3, bool cumulative, double expected)
        {
            ValueEval result = invokeValue(number1, number2, number3, cumulative);
            Assert.IsTrue(result is NumberEval);
            Assert.AreEqual(expected, ((NumberEval)result).NumberValue, 0.00000000000001);
        }

        private static void confirmInvalidError(String number1, String number2, String number3, bool cumulative)
        {
            ValueEval result = invokeValue(number1, number2, number3, cumulative);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorEval.VALUE_INVALID, result);
        }

        private static void confirmNumError(String number1, String number2, String number3, bool cumulative)
        {
            ValueEval result = invokeValue(number1, number2, number3, cumulative);
            Assert.IsTrue(result is ErrorEval);
            Assert.AreEqual(ErrorEval.NUM_ERROR, result);
        }
    }
}