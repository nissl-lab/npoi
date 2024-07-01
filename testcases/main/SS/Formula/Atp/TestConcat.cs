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
    public class TestConcat
    {
        private IWorkbook wb;
        private ISheet sheet;
        private IFormulaEvaluator evaluator;
        private ICell textCell1;
        private ICell textCell2;
        private ICell numericCell1;
        private ICell numericCell2;
        private ICell textCell3;
        private ICell formulaCell;

        private IWorkbook InitWorkbook1()
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet sheet = wb.CreateSheet("Sheet1");

            sheet.CreateRow(0).CreateCell(0).SetCellValue("Currency");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("US Dollar");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("Australian Dollar");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("Chinese Yuan");
            sheet.CreateRow(4).CreateCell(0).SetCellValue("Hong Kong Dollar");
            sheet.CreateRow(5).CreateCell(0).SetCellValue("Israeli Shekel");
            sheet.CreateRow(6).CreateCell(0).SetCellValue("South Korean Won");
            sheet.CreateRow(7).CreateCell(0).SetCellValue("Russian Ruble");
            return wb;
        }

        [SetUp]
        public void Setup() 
        {
            wb = new HSSFWorkbook();
            evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            sheet = wb.CreateSheet("CONCAT");
            IRow row = sheet.CreateRow(0);

            textCell1 = row.CreateCell(0);
            textCell1.SetCellValue("One");

            textCell2 = row.CreateCell(1);
            textCell2.SetCellValue("Two");

            textCell3 = row.CreateCell(2);
            textCell3.SetCellValue("Three");

            numericCell1 = row.CreateCell(3);
            numericCell1.SetCellValue(1);

            numericCell2 = row.CreateCell(4);
            numericCell2.SetCellValue(2);

            formulaCell = row.CreateCell(100, CellType.Formula);
        }

        [Test]
        public void TestConcatWithStrings()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("CONCAT(\"The\",\" \",\"sun\",\" \",\"will\",\" \",\"come\",\" \",\"up\",\" \",\"tomorrow.\")");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("The sun will come up tomorrow.", formulaCell.StringCellValue);
        }

        [Test]
        public void TestConcatWithColumns()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("CONCAT(B:B, C:C)");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("TwoThree", formulaCell.StringCellValue);
        }

        [Test]
        public void TestConcatWithCellRanges()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("CONCAT(A1:C1)");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("OneTwoThree", formulaCell.StringCellValue);
        }

        [Test]
        public void TestConcatWithCellRefs()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("CONCAT(\"ONE\", A1, \"TWO\",B1, \"THREE\",C1, \".\")");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("ONEOneTWOTwoTHREEThree.", formulaCell.StringCellValue);
        }

    }
}
