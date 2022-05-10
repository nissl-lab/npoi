using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestAreas
    {
        [Test]
        public void TestBasic()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ICell cell = wb.CreateSheet().CreateRow(0).CreateCell(0);
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);

            String formulaText = "AREAS(B1)";
            confirmResult(fe, cell, formulaText, 1.0);

            formulaText = "AREAS(B2:D4)";
            confirmResult(fe, cell, formulaText, 1.0);

            formulaText = "AREAS((B2:D4,E5,F6:I9))";
            confirmResult(fe, cell, formulaText, 3.0);

            formulaText = "AREAS((B2:D4,E5,C3,E4))";
            confirmResult(fe, cell, formulaText, 4.0);

            formulaText = "AREAS((I9))";
            confirmResult(fe, cell, formulaText, 1.0);
        }

        private static void confirmResult(HSSFFormulaEvaluator fe, ICell cell, String formulaText, Double expectedResult)
        {
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Numeric,result.CellType);
            Assert.AreEqual(expectedResult, result.NumberValue);
        }
    }
}
