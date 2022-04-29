using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TestCases.XSSF
{
    [TestFixture]
    public class TestXLookupFunction
    {
        [Test]
        public void TestMicrosoftExample1()
        {
            String formulaText = "XLOOKUP(B2,B5:B14,C5:D14)";
            XSSFWorkbook wb = initWorkbook2();
            XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(wb);
            ISheet sheet = wb.GetSheetAt(0);
            IRow row1 = sheet.GetRow(1);
            String col1 = CellReference.ConvertNumToColString(2);
            String col2 = CellReference.ConvertNumToColString(3);
            String cellRef = $"{col1}2:{col2}2";
            sheet.SetArrayFormula(formulaText, CellRangeAddress.ValueOf(cellRef));
            fe.EvaluateAll();
            FileInfo fi = TempFile.CreateTempFile("xlook", ".xlsx");
            using (FileStream file = new FileStream(fi.FullName, FileMode.Open, FileAccess.ReadWrite))
            {
                wb.Write(file, false);
            }
            Assert.AreEqual("Dianne Pugh", row1.GetCell(2).StringCellValue);
            //next assertion fails, cell D2 ends up with Dianne Pugh
            Assert.AreEqual("Finance", row1.GetCell(3).StringCellValue);
        }
        private XSSFWorkbook initWorkbook2()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet();
            SS.Util.Utils.AddRow(sheet, 0, null, "Emp Id", "Employee Name", "Department");
            SS.Util.Utils.AddRow(sheet, 1, null, 8389);
            SS.Util.Utils.AddRow(sheet, 3, null, "Emp Id", "Employee Name", "Department");
            SS.Util.Utils.AddRow(sheet, 4, null, 4390, "Ned Lanning", "Marketing");
            SS.Util.Utils.AddRow(sheet, 5, null, 8604, "Margo Hendrix", "Sales");
            SS.Util.Utils.AddRow(sheet, 6, null, 8389, "Dianne Pugh", "Finance");
            SS.Util.Utils.AddRow(sheet, 7, null, 4937, "Earlene McCarty", "Accounting");
            SS.Util.Utils.AddRow(sheet, 8, null, 8299, "Mia Arnold", "Operation");
            SS.Util.Utils.AddRow(sheet, 9, null, 2643, "Jorge Fellows", "Executive");
            SS.Util.Utils.AddRow(sheet, 10, null, 5243, "Rose Winters", "Sales");
            SS.Util.Utils.AddRow(sheet, 11, null, 9693, "Carmela Hahn", "Finance");
            SS.Util.Utils.AddRow(sheet, 12, null, 1636, "Delia Cochran", "Accounting");
            SS.Util.Utils.AddRow(sheet, 13, null, 6703, "Marguerite Cervantes", "Marketing");
            return wb;
        }
    }
}
