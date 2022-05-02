using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.SS.Formula.Functions
{
    [TestFixture]
    public class TestPercentRank
    {
        //https://support.microsoft.com/en-us/office/percentrank-function-f1b5836c-9619-4847-9fc9-080ec9024442
        [Test]
        public void TestMicrosoftExample1()
        {
            HSSFWorkbook wb = initWorkbook1();
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            confirmNumericResult(fe, cell, "PERCENTRANK(A2:A11,2)", 0.333);
            confirmNumericResult(fe, cell, "PERCENTRANK(A2:A11,8,2)", 0.66);
            confirmNumericResult(fe, cell, "PERCENTRANK(A2:A11,8,4)", 0.6666);
            confirmNumericResult(fe, cell, "PERCENTRANK(A2:A11,4)", 0.555);
            confirmNumericResult(fe, cell, "PERCENTRANK(A2:A11,8)", 0.666);
            confirmNumericResult(fe, cell, "PERCENTRANK(A2:A11,5)", 0.583);

        }

        [Test]
        public void TestErrorCases()
        {
            HSSFWorkbook wb = initWorkbook1();
            HSSFFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            confirmErrorResult(fe, cell, "PERCENTRANK(A2:A11,0)", FormulaError.NA);
            confirmErrorResult(fe, cell, "PERCENTRANK(A2:A11,100)", FormulaError.NA);
            confirmErrorResult(fe, cell, "PERCENTRANK(B2:B11,100)", FormulaError.NUM);

        }

        private HSSFWorkbook initWorkbook1()
        {
            HSSFWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, "Data");
            SS.Util.Utils.AddRow(sheet, 1, 13);
            SS.Util.Utils.AddRow(sheet, 2, 12);
            SS.Util.Utils.AddRow(sheet, 3, 11);
            SS.Util.Utils.AddRow(sheet, 4, 8);
            SS.Util.Utils.AddRow(sheet, 5, 4);
            SS.Util.Utils.AddRow(sheet, 6, 3);
            SS.Util.Utils.AddRow(sheet, 7, 2);
            SS.Util.Utils.AddRow(sheet, 8, 1);
            SS.Util.Utils.AddRow(sheet, 9, 1);
            SS.Util.Utils.AddRow(sheet, 10, 1);
            return wb;
        }

        private static void confirmNumericResult(HSSFFormulaEvaluator fe, ICell cell, String formulaText, double expectedResult)
        {
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(CellType.Numeric, result.CellType);
            Assert.AreEqual(expectedResult, result.NumberValue, 0.0001);
        }

        private static void confirmErrorResult(HSSFFormulaEvaluator fe, ICell cell, String formulaText, FormulaError expectedResult)
        {
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(expectedResult.Code, result.ErrorValue);
        }
    }
}
