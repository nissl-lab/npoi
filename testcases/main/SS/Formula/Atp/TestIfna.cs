using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Atp
{
    [TestFixture]
    public class TestIfna
    {
        IWorkbook wb;
        ICell cell;
        IFormulaEvaluator fe;
        [SetUp]
        public void Setup()
        {
            wb = new HSSFWorkbook();
            cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            fe = new HSSFFormulaEvaluator(wb);
        }
        [Test]
        public void testNumbericArgsWorkCorrectly()
        {
            ConfirmResult(fe, cell, "IFNA(-1,42)", new CellValue(-1.0));
            ConfirmResult(fe, cell, "IFNA(NA(),42)", new CellValue(42.0));
        }

        [Test]
        public void testStringArgsWorkCorrectly()
        {
            ConfirmResult(fe, cell, "IFNA(\"a1\",\"a2\")", new CellValue("a1"));
            ConfirmResult(fe, cell, "IFNA(NA(),\"a2\")", new CellValue("a2"));
        }

        [Test]
        public void testUsageErrorsThrowErrors()
        {
            ConfirmError(fe, cell, "IFNA(1)", ErrorEval.VALUE_INVALID);
            ConfirmError(fe, cell, "IFNA(1,2,3)", ErrorEval.VALUE_INVALID);
        }

        [Test]
        public void testErrorInArgSelectsNAResult()
        {
            ConfirmError(fe, cell, "IFNA(1/0,42)", ErrorEval.DIV_ZERO);
        }

        [Test]
        public void testErrorFromNAArgPassesThrough()
        {
            ConfirmError(fe, cell, "IFNA(NA(),1/0)", ErrorEval.DIV_ZERO);
        }

        [Test]
        public void testNaArgNotEvaledIfUnneeded()
        {
            ConfirmResult(fe, cell, "IFNA(42,1/0)", new CellValue(42.0));
        }

        private static void ConfirmResult(IFormulaEvaluator fe, ICell cell, String formulaText, CellValue expectedResult)
        {
            fe.DebugEvaluationOutputForNextEval = true;
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(expectedResult.CellType, result.CellType, "Testing result type for: " + formulaText);
            Assert.AreEqual(expectedResult.FormatAsString(), result.FormatAsString(), "Testing result for: " + formulaText);
        }

        private static void ConfirmError(IFormulaEvaluator fe, ICell cell, String formulaText,ErrorEval expectedError)
        {
            fe.DebugEvaluationOutputForNextEval = true;
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Error, result.CellType, "Testing result type for: " + formulaText);
            Assert.AreEqual(expectedError.ErrorString, result.FormatAsString(), "Testing error type for: " + formulaText);
        }
    }
}
