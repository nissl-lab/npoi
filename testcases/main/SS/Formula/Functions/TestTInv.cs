using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestTInv
    {
        [Test]
        public void TestBasic()
        {
            HSSFWorkbook wb = new HSSFWorkbook();

            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            wb.GetSheetAt(0).GetRow(0).CreateCell(1).SetCellValue(0.75);
            wb.GetSheetAt(0).CreateRow(1).CreateCell(1).SetCellValue(2);

            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            String formulaText = "T.INV(0.75,2)";
            confirmResult(fe, cell, formulaText, 0.816496581);

            formulaText = "T.INV(0.99,100)*126+36/2";
            confirmResult(fe, cell, formulaText, 315.8913881);

            formulaText = "T.INV(B1,B2)";
            confirmResult(fe, cell, formulaText, 0.816496581);

            formulaText = "T.INV(0.75, 2, 1)";
            confirmNumError(fe, cell, formulaText);

            formulaText = "T.INV(B1,B2,C3)";
            confirmNumError(fe, cell, formulaText);

        }

        private static void confirmResult(HSSFFormulaEvaluator fe, ICell cell, String formulaText, Double expectedResult)
        {
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Numeric,result.CellType);
            Assert.AreEqual(expectedResult, result.NumberValue, 0.0000001);
        }

        private static void confirmNumError(HSSFFormulaEvaluator fe, ICell cell, String formulaText)
        {
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Error, result.CellType);
        }
    }
}
