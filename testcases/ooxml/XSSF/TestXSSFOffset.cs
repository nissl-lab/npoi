using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCases.XSSF
{
    [TestFixture]
    class TestXSSFOffset
    {
        [Test]
        public void TestOffsetWithEmpty23Arguments()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ICell cell = workbook.CreateSheet().CreateRow(0).CreateCell(0);
            cell.SetCellFormula("OFFSET(B1,,)");
            String value = "EXPECTED_VALUE";
            ICell valueCell = cell.Row.CreateCell(1);
            valueCell.SetCellValue(value);
            workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
            Assert.AreEqual(CellType.String, cell.CachedFormulaResultType);
            Assert.AreEqual(value, cell.StringCellValue);
        }
    }
}
