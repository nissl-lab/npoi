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
    public class TestTextJoinFunction
    {
        private IWorkbook wb;
        private ISheet sheet;
        private IFormulaEvaluator evaluator;
        private ICell textCell1;
        private ICell textCell2;
        private ICell numericCell1;
        private ICell numericCell2;
        private ICell blankCell;
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
        private static void ConfirmResult(IFormulaEvaluator fe, ICell cell, String formulaText, String expectedResult)
        {
            cell.SetCellFormula(formulaText);
            fe.NotifyUpdateCell(cell);
            CellValue result = fe.Evaluate(cell);
            Assert.AreEqual(result.CellType, CellType.String);
            Assert.AreEqual(expectedResult, result.StringValue);
        }
        [SetUp]
        public void SetUp()
        {
            wb = new HSSFWorkbook();
            evaluator = wb.GetCreationHelper().CreateFormulaEvaluator();

            sheet = wb.CreateSheet("TextJoin");
            IRow row = sheet.CreateRow(0);

            textCell1 = row.CreateCell(0);
            textCell1.SetCellValue("One");

            textCell2 = row.CreateCell(1);
            textCell2.SetCellValue("Two");

            blankCell = row.CreateCell(2);
            blankCell.SetBlank();

            numericCell1 = row.CreateCell(3);
            numericCell1.SetCellValue(1);

            numericCell2 = row.CreateCell(4);
            numericCell2.SetCellValue(2);

            formulaCell = row.CreateCell(100, CellType.Formula);
        }
        [Test]
        public void TestJoinSingleLiteralText()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", true, \"Text\")");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("Text", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinMultipleLiteralText()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", true, \"One\", \"Two\", \"Three\")");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("One,Two,Three", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinLiteralTextAndNumber()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", true, \"Text\", 1)");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("Text,1", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinEmptyStringIncludeEmpty()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", false, \"A\", \"\", \"B\")");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("A,,B", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinEmptyStringIgnoreEmpty()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", true, \"A\", \"\", \"B\")");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("A,B", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinEmptyStringsIncludeEmpty()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", false, \"\", \"\")");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(",", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinEmptyStringsIgnoreEmpty()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", true, \"\", \"\")");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinTextCellValues()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", true, A1, B1)");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("One,Two", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinNumericCellValues()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", true, D1, E1)");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("1,2", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinBlankCellIncludeEmpty()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", false, A1, C1, B1)");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("One,,Two", formulaCell.StringCellValue);
        }

        [Test]
        public void TestJoinBlankCellIgnoreEmpty()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", true, A1, C1, B1)");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual("One,Two", formulaCell.StringCellValue);
        }

        [Test]
        public void TestNoTextArgument()
        {
            evaluator.ClearAllCachedResultValues();
            formulaCell.SetCellFormula("TEXTJOIN(\",\", true)");
            evaluator.EvaluateFormulaCell(formulaCell);
            Assert.AreEqual(CellType.Error, formulaCell.CachedFormulaResultType);
            Assert.AreEqual(ErrorEval.VALUE_INVALID.ErrorCode, formulaCell.ErrorCellValue);
        }

        //https://support.microsoft.com/en-us/office/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c
        [Test]
        public void TestMicrosoftExample1()
        {
            IWorkbook wb = InitWorkbook1();

            IFormulaEvaluator fe = new HSSFFormulaEvaluator(wb);
            ICell cell = wb.GetSheetAt(0).GetRow(0).CreateCell(100);
            ConfirmResult(fe, cell, "TEXTJOIN(\", \", TRUE, A2:A8)",
                        "US Dollar, Australian Dollar, Chinese Yuan, Hong Kong Dollar, Israeli Shekel, South Korean Won, Russian Ruble");

        }

    }
}
